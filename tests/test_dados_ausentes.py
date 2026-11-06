"""Linhas não preenchidas e leito-dia sem crônicos inconsistente. Números sintéticos."""

import openpyxl

from extracao_producao import ler_producao
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


def test_linha_vazia_com_formula_zerada_ao_lado_nao_vira_zero():
    outro = BLOCO + "__nascidos"
    dados, _, _ = sanitizar(
        {BLOCO: 0.0, FILHO: 0.0, outro: 0.0},
        {BLOCO: [None] * 2, FILHO: [None] * 2, outro: [0.0, 0.0]},
    )

    assert dados[FILHO] == SEM_DADOS and dados[outro] == SEM_DADOS


def _planilha_com_texto(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "UTI SET"
    linhas = [
        ["UTI - SETEMBRO/2026"],
        ["DATAS", 1, 2, "TOTAL", "REALIZADO"],
        ["PACIENTE DIA", 8, "1O", 8],
        [],
        ["TAXA EXEMPLO", None, None, "TOTAL", 0.5],
        ["N. DE PACIENTES-DIA", 4, 4, "#VALUE!"],
        ["N. DE LEITO-DIAS", 6, 6, 12],
    ]
    for linha in linhas:
        ws.append(linha)
    caminho = tmp_path / "producao.xlsx"
    wb.save(caminho)
    return caminho


def test_texto_em_celula_de_dia_e_informado(tmp_path):
    _, _, _, meta = ler_producao(_planilha_com_texto(tmp_path), "SETEMBRO")

    assert meta["textos"] == [{"celula": "C3", "rotulo": "PACIENTE DIA", "valor": "1O"}]


def test_erro_de_formula_no_total_nao_vira_titulo_de_bloco(tmp_path):
    totais, _, _, _ = ler_producao(_planilha_com_texto(tmp_path), "SETEMBRO")

    assert totais["taxa_exemplo__n_de_leito_dias"] == 12
    assert "n_de_pacientes_dia__n_de_leito_dias" not in totais


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
