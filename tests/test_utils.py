from datetime import date

import pytest

from utils import eh_numero, normalizar, resolver_ano, resolver_mes, resolver_periodo, sem_acento


@pytest.mark.parametrize("entrada", ["setembro", "SETEMBRO", "Setembro", "set", "SET", "Set.", " set ", "9", "09", 9])
def test_resolver_mes_aceita_variantes_de_setembro(entrada):
    assert resolver_mes(entrada) == "SETEMBRO"


@pytest.mark.parametrize("entrada", ["março", "MARÇO", "marco", "mar", "3", "03"])
def test_resolver_mes_aceita_marco_com_e_sem_cedilha(entrada):
    assert resolver_mes(entrada) == "MARCO"


@pytest.mark.parametrize("entrada", ["13", "0", "00", "", "xyz", "se", "setem", "-1", "1.5", "set2026"])
def test_resolver_mes_rejeita_entradas_invalidas(entrada):
    with pytest.raises(ValueError, match="Mês inválido"):
        resolver_mes(entrada)


@pytest.mark.parametrize("entrada,esperado", [("2026", 2026), (2026, 2026), ("26", 2026), (" 2025 ", 2025), ("00", 2000)])
def test_resolver_ano(entrada, esperado):
    assert resolver_ano(entrada) == esperado


@pytest.mark.parametrize("entrada", ["abc", "202", "20266", "26.0", "1999", "2101", "0026"])
def test_resolver_ano_rejeita_entradas_invalidas(entrada):
    with pytest.raises(ValueError, match="Ano"):
        resolver_ano(entrada)


@pytest.mark.parametrize("vazio", [None, ""])
def test_resolver_ano_padrao_e_o_ano_atual(vazio):
    assert resolver_ano(vazio) == date.today().year


@pytest.mark.parametrize("mes,ano,n_dias", [
    ("fev", 2024, 29),  # bissexto
    ("fev", 2026, 28),
    ("jul", 2026, 31),
    ("ago", 2026, 31),
    ("set", 2026, 30),
    ("dez", 2026, 31),
])
def test_resolver_periodo_calcula_dias_do_mes(mes, ano, n_dias):
    assert resolver_periodo(mes, ano).n_dias == n_dias


@pytest.mark.parametrize("entrada", ["set/2026", "09-2026", "setembro 2026", "SET/26", "9/26"])
def test_resolver_periodo_aceita_mes_e_ano_juntos(entrada):
    periodo = resolver_periodo(entrada)
    assert (periodo.mes, periodo.ano, periodo.numero, periodo.n_dias) == ("SETEMBRO", 2026, 9, 30)


def test_resolver_periodo_ano_explicito_vence_e_nao_divide_o_mes():
    assert resolver_periodo("setembro", "2025") == ("SETEMBRO", 2025, 9, 30)


@pytest.mark.parametrize("entrada", ["set/20/26", "2026-09", "13/2026", "set/abc"])
def test_resolver_periodo_rejeita_formatos_invalidos(entrada):
    with pytest.raises(ValueError):
        resolver_periodo(entrada)


def test_sem_acento_e_normalizar():
    assert sem_acento("Produção Diária") == "Producao Diaria"
    assert sem_acento("CVC – DIA") == "CVC - DIA"  # travessão vira hífen em vez de sumir
    assert normalizar("  n. de   saídas ") == "N. DE SAIDAS"


@pytest.mark.parametrize("valor,esperado", [
    (0, True), (1.5, True), (-3, True),
    (True, False), (False, False), (None, False), ("12", False), ("[Sem Dados]", False),
    (float("nan"), False), (float("inf"), False),
])
def test_eh_numero(valor, esperado):
    assert eh_numero(valor) is esperado
