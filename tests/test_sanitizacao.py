"""Regras de sanitização com números sintéticos (não são dados do hospital)."""

import pytest

from sanitizacao import PENDENTE_INVALIDO, sanitizar

OCUP = "taxa_de_ocupacao_global"
PERM = "media_de_permanencia"
ADESAO = "adesao_as_metas_de_identificacao_do_paciente"
ADESAO_NUM = ADESAO + "__utilizar_a_pulseira_de_identificacao_e_a_dupla_chegagem_de_medicamentos"


def base(**extra):
    """Mês fictício: 100 pacientes-dia, 120 leitos-dia, 10 saídas, 20 pacientes-dia crônicos."""
    dados = {
        "paciente_dia_global": 100.0,
        "paciente_dia_cronicos": 20.0,
        "leito_dia_global": 120.0,
        "leito_dia_sem_os_cronicos": 90.0,
        "saidas_hospitalares": 10.0,
    }
    dados.update(extra)
    return dados


def niveis(log):
    return [nivel for nivel, _ in log]


# --- Censo x Produção --------------------------------------------------------------------------

def test_censo_divergente_adota_producao_com_info():
    dados, log, correcoes = sanitizar(base(censo_paciente_dia=95.0, censo_total_saidas_geral=10.0))
    assert dados["censo_paciente_dia"] == 100.0
    assert dados["censo_total_saidas_geral"] == 10.0  # igual: intocado
    assert niveis(log) == ["INFO"]
    assert correcoes == []  # conciliação não gera write-back


def test_censo_pendente_nao_e_conciliado():
    dados, log, _ = sanitizar(base(censo_paciente_dia="[Sem Dados]"))
    assert dados["censo_paciente_dia"] == "[Sem Dados]"
    assert log == []


# --- Denominadores repetidos dentro dos blocos -------------------------------------------------

def test_paciente_dia_do_bloco_recebe_o_global():
    chave = PERM + "__n_de_paciente_dia"
    dados, log, correcoes = sanitizar(base(**{chave: 10.0}))
    assert dados[chave] == 100.0
    assert niveis(log) == ["CORREÇÃO"]
    assert correcoes == [{"chave": chave, "novo": 100.0, "acao": ("ref", "paciente_dia_global")}]


def test_paciente_dia_sem_cronicos_e_global_menos_cronicos():
    chave = "taxa_de_ocupacao_sem_os_cronicos__n_de_pacientes_dia"
    dados, _, correcoes = sanitizar(base(**{chave: 1.0}))
    assert dados[chave] == 80.0
    assert correcoes[0]["acao"] == ("sub", "paciente_dia_global", "paciente_dia_cronicos")


def test_leito_dia_do_bloco_recebe_o_global():
    chave = OCUP + "__n_de_leito_dias"
    dados, _, correcoes = sanitizar(base(**{chave: 12.0}))
    assert dados[chave] == 120.0
    assert correcoes[0]["acao"] == ("ref", "leito_dia_global")


def test_correcao_vira_soma_da_linha_quando_os_dias_batem():
    chave = PERM + "__n_de_paciente_dia"
    series = {chave: [50.0, 30.0, None, 20.0]}  # soma 100 = global: a fórmula certa é SOMA da própria linha
    _, _, correcoes = sanitizar(base(**{chave: 7.0}), series)
    assert correcoes[0]["acao"] == ("soma_linha",)


def test_diferenca_dentro_da_tolerancia_nao_corrige():
    chave = PERM + "__n_de_paciente_dia"
    dados, log, _ = sanitizar(base(**{chave: 100.4}))
    assert dados[chave] == 100.4
    assert log == []


# --- Razões (taxa = numerador / denominador) ---------------------------------------------------

def razao_ocupacao(taxa):
    return base(**{
        OCUP: taxa,
        OCUP + "__n_de_pacientes_dia": 100.0,
        OCUP + "__n_de_leito_dias": 120.0,
    })


def test_razao_quebrada_e_recalculada():
    dados, log, correcoes = sanitizar(razao_ocupacao(0.5))
    assert dados[OCUP] == pytest.approx(100 / 120)
    assert "CORREÇÃO" in niveis(log)
    assert {"chave": OCUP, "novo": pytest.approx(100 / 120),
            "acao": ("div", OCUP + "__n_de_pacientes_dia", OCUP + "__n_de_leito_dias")} in correcoes


def test_razao_correta_nao_e_alterada():
    dados, log, correcoes = sanitizar(razao_ocupacao(100 / 120))
    assert log == [] and correcoes == []


def test_razao_com_denominador_zero_e_ignorada():
    dados = base(**{PERM: 3.0, PERM + "__n_de_paciente_dia": 100.0, PERM + "__n_de_saidas": 0.0})
    saida, log, _ = sanitizar(dados)
    assert saida[PERM] == 3.0
    assert log == []


def test_media_de_permanencia_pode_passar_de_1():
    """Média não é proporção: 8,5 dias é válido e não deve ser marcado como fora de 0–100%."""
    dados = base(**{PERM: 10.0, PERM + "__n_de_paciente_dia": 100.0, PERM + "__n_de_saidas": 10.0})
    saida, log, _ = sanitizar(dados)
    assert saida[PERM] == 10.0
    assert log == []


# --- Proporções fora de 0–100% -----------------------------------------------------------------

def test_proporcao_fora_da_faixa_recalculada_pelos_dois_filhos():
    chave = "incidencia_de_queda_de_paciente"
    dados = base(**{chave: 4.0, chave + "__n_de_quedas": 2.0, chave + "__n_de_paciente_dia": 100.0})
    saida, log, correcoes = sanitizar(dados)
    assert saida[chave] == pytest.approx(0.02)
    assert correcoes[-1]["acao"] == ("div", chave + "__n_de_quedas", chave + "__n_de_paciente_dia")


def test_proporcao_nao_e_dividida_por_100():
    """2,57 numa taxa de 100% viraria 2,57% se dividido por 100: tem de recalcular ou marcar pendente."""
    dados = base(**{ADESAO: 2.57, ADESAO_NUM: 100.0})
    saida, _, correcoes = sanitizar(dados)
    assert saida[ADESAO] == pytest.approx(1.0)  # recálculo especial: numerador / paciente-dia global
    assert correcoes[-1]["acao"] == ("div", ADESAO_NUM, "paciente_dia_global")


def test_proporcao_sem_como_recalcular_vira_pendente():
    chave = "indice_de_ulcera_por_pressao"
    saida, log, correcoes = sanitizar(base(**{chave: 1.7}))
    assert saida[chave] == PENDENTE_INVALIDO
    assert niveis(log) == ["WARNING"]
    assert correcoes == []


def test_proporcao_recalculada_ainda_fora_da_faixa_vira_pendente():
    chave = "incidencia_de_flebite"
    dados = base(**{chave: 5.0, chave + "__n_de_flebites": 300.0, chave + "__n_de_paciente_dia_com_cvp": 100.0})
    saida, log, _ = sanitizar(dados)
    assert saida[chave] == PENDENTE_INVALIDO
    assert "WARNING" in niveis(log)


@pytest.mark.parametrize("chave", ["taxa_de_ocupacao_global__n_de_leito_dias", "n_de_nascidos_vivos"])
def test_linhas_de_contagem_nao_sao_tratadas_como_proporcao(chave):
    saida, log, _ = sanitizar({chave: 250.0})
    assert saida[chave] == 250.0
    assert log == []


# --- Contrato geral ----------------------------------------------------------------------------

def test_sanitizar_nao_altera_o_dicionario_recebido():
    entrada = razao_ocupacao(0.5)
    copia = dict(entrada)
    sanitizar(entrada)
    assert entrada == copia


def test_cenario_completo_de_um_mes_ficticio():
    """Bloco de permanência com denominador errado e taxa calculada sobre ele (padrão visto na planilha)."""
    dados = base(**{
        "censo_paciente_dia": 97.0,
        PERM: 1.0,
        PERM + "__n_de_paciente_dia": 10.0,  # fórmula somando a linha errada
        PERM + "__n_de_saidas": 10.0,
        ADESAO: 2.57,
        ADESAO_NUM: 100.0,
    })
    saida, log, correcoes = sanitizar(dados)
    assert saida["censo_paciente_dia"] == 100.0
    assert saida[PERM + "__n_de_paciente_dia"] == 100.0
    assert saida[PERM] == pytest.approx(10.0)  # recalculada DEPOIS de corrigir o denominador
    assert saida[ADESAO] == pytest.approx(1.0)
    assert niveis(log) == ["INFO", "CORREÇÃO", "CORREÇÃO", "CORREÇÃO"]
    assert [c["chave"] for c in correcoes] == [PERM + "__n_de_paciente_dia", PERM, ADESAO]
