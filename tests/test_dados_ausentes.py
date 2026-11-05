"""Linhas não preenchidas e leito-dia sem crônicos inconsistente. Números sintéticos."""

from sanitizacao import SEM_DADOS, sanitizar
from validacao import _conferir_leito_sem_cronicos

BLOCO = "taxa_exemplo"
FILHO = BLOCO + "__adesao"


def test_bloco_com_todos_os_dias_vazios_nao_vira_zero():
    dados, log, correcoes = sanitizar({BLOCO: 0.0, FILHO: 0.0}, {BLOCO: [None] * 3, FILHO: [None] * 3})

    assert dados[BLOCO] == SEM_DADOS and dados[FILHO] == SEM_DADOS
    assert [nivel for nivel, _ in log] == ["WARNING"]
    assert correcoes == []


def test_zero_lancado_nos_dias_continua_zero():
    dados, log, _ = sanitizar({BLOCO: 0.0, FILHO: 0.0}, {BLOCO: [None] * 3, FILHO: [0.0, 0.0, 0.0]})

    assert dados[BLOCO] == 0.0 and dados[FILHO] == 0.0
    assert log == []


def test_bloco_parcialmente_vazio_nao_e_marcado():
    outro = BLOCO + "__total"
    dados, _, _ = sanitizar(
        {BLOCO: 0.0, FILHO: 0.0, outro: 5.0},
        {BLOCO: [None] * 2, FILHO: [None] * 2, outro: [2.0, 3.0]},
    )

    assert dados[FILHO] == 0.0


def _series(ultimo_sem_cronicos):
    return {
        "leito_dia_global": [10.0, 10.0],
        "paciente_dia_cronicos": [4.0, 4.0],
        "leito_dia_sem_os_cronicos": [6.0, ultimo_sem_cronicos],
    }


def test_leito_sem_cronicos_zerado_gera_aviso_com_impacto():
    dados = {
        "leito_dia_sem_os_cronicos": 6.0,
        "taxa_de_ocupacao_sem_os_cronicos": 0.5,
        "taxa_de_ocupacao_sem_os_cronicos__n_de_pacientes_dia": 3.0,
    }
    problemas = _conferir_leito_sem_cronicos(dados, _series(0.0))

    assert [nivel for nivel, _ in problemas] == ["AVISO"]
    assert "dia 2: 0 (esperado 6)" in problemas[0][1]
    assert "esperado: 12" in problemas[0][1]
    assert "de 50,00% para 25,00%." in problemas[0][1]


def test_leito_sem_cronicos_consistente_nao_avisa():
    assert _conferir_leito_sem_cronicos({"leito_dia_sem_os_cronicos": 12.0}, _series(6.0)) == []
