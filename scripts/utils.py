"""Funções comuns: normalização de texto, teste numérico e interpretação de mês/ano."""

import calendar
import math
import re
import sys
import unicodedata
from collections import namedtuple
from datetime import date

# Nome completo (sem acento) -> (nome, abreviação, número com dois dígitos). A ordem é a do calendário.
MESES = {
    "JANEIRO": ("JANEIRO", "JAN", "01"),
    "FEVEREIRO": ("FEVEREIRO", "FEV", "02"),
    "MARCO": ("MARCO", "MAR", "03"),
    "ABRIL": ("ABRIL", "ABR", "04"),
    "MAIO": ("MAIO", "MAI", "05"),
    "JUNHO": ("JUNHO", "JUN", "06"),
    "JULHO": ("JULHO", "JUL", "07"),
    "AGOSTO": ("AGOSTO", "AGO", "08"),
    "SETEMBRO": ("SETEMBRO", "SET", "09"),
    "OUTUBRO": ("OUTUBRO", "OUT", "10"),
    "NOVEMBRO": ("NOVEMBRO", "NOV", "11"),
    "DEZEMBRO": ("DEZEMBRO", "DEZ", "12"),
}
ANO_MIN, ANO_MAX = 2000, 2100
# Marca das cópias gravadas por corrigir_planilha (<nome>_Corrigida_<MES>.xlsx); nunca são lidas como fonte.
SUFIXO_CORRIGIDA = "_Corrigida"

Periodo = namedtuple("Periodo", "mes ano numero n_dias")
Periodo.__doc__ = "Mês de referência: mes = chave de MESES, ano = int, numero = 1–12, n_dias = dias do mês."


def sem_acento(texto):
    """'Produção – UTI' -> 'Producao - UTI'. Travessões viram hífen antes de descartar o não-ASCII."""
    texto = str(texto).replace("–", "-").replace("—", "-")
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()


def normalizar(texto):
    """Maiúsculas, sem acentos, sem espaços nas pontas e sem espaços duplicados."""
    return " ".join(sem_acento(texto).upper().split())


def eh_numero(valor):
    """True para int/float finitos. bool, NaN, infinito e marcadores de pendência (str) não contam."""
    if isinstance(valor, bool) or not isinstance(valor, (int, float)):
        return False
    return math.isfinite(valor)


def resolver_mes(mes):
    """Aceita 'SETEMBRO', 'setembro', 'Set', 'set', '9' ou '09' e devolve a chave de MESES ('SETEMBRO')."""
    alvo = normalizar(mes).rstrip(".")
    if alvo.isdigit():
        n = int(alvo)
        if 1 <= n <= 12:
            return list(MESES)[n - 1]
    else:
        for chave, (nome, abrev, _) in MESES.items():
            if alvo in (nome, abrev):
                return chave
    raise ValueError(f"Mês inválido: {mes!r}. Use o nome ('setembro'), a abreviação ('set') ou o número (9).")


def resolver_ano(ano=None):
    """'2026' ou 2026 -> 2026; '26' -> 2026; None -> ano atual. Fora de 2000–2100 é erro."""
    if ano is None or str(ano).strip() == "":
        return date.today().year
    texto = str(ano).strip()
    if not re.fullmatch(r"\d{2}|\d{4}", texto):
        raise ValueError(f"Ano inválido: {ano!r}. Use quatro dígitos (2026) ou dois (26).")
    valor = int(texto) + (2000 if len(texto) == 2 else 0)
    if not ANO_MIN <= valor <= ANO_MAX:
        raise ValueError(f"Ano fora do intervalo {ANO_MIN}–{ANO_MAX}: {ano!r}.")
    return valor


def resolver_periodo(mes, ano=None):
    """Interpreta mês e ano da linha de comando e devolve um Periodo.

    O mês pode vir junto com o ano quando `ano` é omitido: 'set/2026', '09-2026', 'setembro 26'.
    Levanta ValueError com mensagem pronta para o usuário em qualquer entrada inválida.
    """
    texto = str(mes).strip()
    if ano is None:
        partes = re.split(r"[/\-\s]+", texto)
        if len(partes) == 2 and partes[1].isdigit():
            texto, ano = partes
        elif len(partes) > 2:
            raise ValueError(f"Período inválido: {mes!r}. Exemplos: 'set 2026', 'setembro/2026', '09-2026'.")
    chave = resolver_mes(texto)
    ano = resolver_ano(ano)
    numero = list(MESES).index(chave) + 1
    return Periodo(chave, ano, numero, calendar.monthrange(ano, numero)[1])


def periodo_da_cli(argv, uso):
    """Para os scripts avulsos: lê MES [ANO] de argv e encerra com a mensagem de uso se for inválido."""
    if not argv:
        sys.exit(f"Uso: {uso}")
    try:
        return resolver_periodo(argv[0], argv[1] if len(argv) > 1 else None)
    except ValueError as e:
        sys.exit(f"{e}\nUso: {uso}")
