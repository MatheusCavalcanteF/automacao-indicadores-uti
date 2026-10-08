"""Log de execução: terminal (mesmas mensagens de sempre) + arquivo de auditoria por execução.

    log = log_execucao.obter(__name__)      # em cada módulo
    log.info("[CORREÇÃO] ...")

main.py chama configurar() no início, que cria relatorios_gerados/execucao_<MES>_<ANO>_<AAAAMMDD_HHMMSS>.log.
Terminal: só a mensagem, a partir de INFO. Arquivo: data/hora, nível, módulo e também o nível DEBUG
(quem rodou, máquina, comando, hash dos arquivos de entrada).
Sem configurar() (testes, import avulso), nada é impresso abaixo de WARNING.
"""

import logging
import sys
from datetime import datetime
from pathlib import Path

RAIZ_LOG = "uti"
FORMATO_ARQUIVO = "%(asctime)s %(levelname)-8s %(name)-26s %(message)s"
FORMATO_DATA = "%Y-%m-%d %H:%M:%S"


class _Terminal(logging.StreamHandler):
    """Escreve sempre no sys.stdout atual, mesmo que ele seja trocado depois de configurar()."""

    def __init__(self):
        super().__init__(sys.stdout)

    @property
    def stream(self):
        return sys.stdout

    @stream.setter
    def stream(self, _):
        pass


def obter(nome):
    """Logger do módulo, filho de 'uti' (não mistura com logs de bibliotecas como pdfminer)."""
    return logging.getLogger(f"{RAIZ_LOG}.{nome}")


def nome_do_arquivo(periodo, agora=None):
    return f"execucao_{periodo.mes}_{periodo.ano}_{(agora or datetime.now()):%Y%m%d_%H%M%S}.log"


def configurar(pasta=None, periodo=None, agora=None):
    """Liga a saída no terminal e, se `pasta` e `periodo` forem dados, o arquivo de auditoria.

    Pode ser chamado de novo: os handlers anteriores são fechados antes. Retorna o Path do log ou None.
    """
    encerrar()
    raiz = logging.getLogger(RAIZ_LOG)
    raiz.setLevel(logging.DEBUG)
    raiz.propagate = False

    try:  # console do Windows em cp1252 não derruba a execução por causa de um "≠" ou "—"
        sys.stdout.reconfigure(errors="replace")
    except (AttributeError, ValueError):
        pass
    tela = _Terminal()
    tela.setLevel(logging.INFO)
    tela.setFormatter(logging.Formatter("%(message)s"))
    raiz.addHandler(tela)

    if pasta is None or periodo is None:
        return None
    caminho = Path(pasta) / nome_do_arquivo(periodo, agora)
    caminho.parent.mkdir(parents=True, exist_ok=True)
    arquivo = logging.FileHandler(caminho, mode="a", encoding="utf-8")
    arquivo.setLevel(logging.DEBUG)
    arquivo.setFormatter(logging.Formatter(FORMATO_ARQUIVO, FORMATO_DATA))
    raiz.addHandler(arquivo)
    return caminho


def encerrar():
    """Fecha e remove os handlers (libera o arquivo de log no Windows)."""
    raiz = logging.getLogger(RAIZ_LOG)
    for handler in list(raiz.handlers):
        raiz.removeHandler(handler)
        handler.close()
