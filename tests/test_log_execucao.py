"""Log de auditoria: arquivo por execução em paralelo à saída no terminal."""

import re
from datetime import datetime

import pytest

import log_execucao
from utils import resolver_periodo

AGORA = datetime(2026, 10, 8, 14, 30, 0)


@pytest.fixture
def log_configurado(tmp_path):
    caminho = log_execucao.configurar(tmp_path / "relatorios_gerados", resolver_periodo("set", 2026), agora=AGORA)
    yield caminho
    log_execucao.encerrar()


def test_nome_do_arquivo_segue_o_padrao(log_configurado):
    assert log_configurado.name == "execucao_SETEMBRO_2026_20261008_143000.log"
    assert log_configurado.parent.name == "relatorios_gerados"  # pasta criada se não existir


def test_mensagens_vao_para_terminal_e_arquivo(log_configurado, capsys):
    log = log_execucao.obter("teste")
    log.debug("hash do arquivo de entrada")
    log.info("[CORREÇÃO] taxa recalculada")
    log.warning("[WARNING] fonte ausente")
    log.error("[ERRO] valor inválido — não recalculável ≠ esperado")
    log_execucao.encerrar()  # fecha o arquivo antes de ler

    terminal = capsys.readouterr().out
    assert terminal.splitlines() == [
        "[CORREÇÃO] taxa recalculada",
        "[WARNING] fonte ausente",
        "[ERRO] valor inválido — não recalculável ≠ esperado",
    ]  # sem data/nível no terminal e sem DEBUG

    linhas = log_configurado.read_text(encoding="utf-8").splitlines()
    assert len(linhas) == 4
    assert "DEBUG" in linhas[0] and "hash do arquivo de entrada" in linhas[0]
    assert "INFO" in linhas[1] and "uti.teste" in linhas[1]
    assert "WARNING" in linhas[2] and "ERROR" in linhas[3]
    assert "≠" in linhas[3]  # UTF-8 preservado
    assert all(re.match(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} ", linha) for linha in linhas)  # data/hora


def test_reconfigurar_nao_duplica_handlers(tmp_path, capsys):
    periodo = resolver_periodo("jul", 2026)
    log_execucao.configurar(tmp_path, periodo, agora=AGORA)
    caminho = log_execucao.configurar(tmp_path, periodo, agora=AGORA)
    log_execucao.obter("teste").info("uma vez só")
    log_execucao.encerrar()
    assert capsys.readouterr().out.count("uma vez só") == 1
    assert caminho.read_text(encoding="utf-8").count("uma vez só") == 1


def test_sem_configurar_nao_cria_arquivo_nem_imprime_info(tmp_path, capsys):
    log_execucao.encerrar()
    log_execucao.obter("teste").info("silencioso")
    assert "silencioso" not in capsys.readouterr().out
    assert not list(tmp_path.iterdir())
