"""Conferências de consistência dos dados antes de gerar o relatório.

Cada problema é uma tupla (nivel, mensagem). Nível "ERRO" = número que não deve ir para o
relatório sem revisão; "AVISO" = divergência a conferir.
"""

import calendar

from localizar_arquivos import MESES, resolver_mes, _sem_acento
from sanitizacao import PREFIXOS_PROPORCAO, RAZOES

# Divergência entre fontes que ainda merece conferência humana (Censo x Produção é tratado na sanitização).
CRUZADOS = [
    ("saídas: produção x bloco de saídas", "saidas_hospitalares", "saidas_hospitalares__n_de_saidas"),
    ("nascidos vivos: produção x ODT (maternidade)", "n_de_nascidos_vivos", "teste_orelhinha"),
]


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def validar(dados, series, titulo_aba, mes, ano):
    problemas = []

    # 1) Aba corresponde ao mês/ano pedidos.
    mes_chave = resolver_mes(mes)
    titulo = _sem_acento(titulo_aba).upper()
    if not any(a in titulo for a in MESES[mes_chave][:2]):
        problemas.append(("ERRO", f"título da aba ('{titulo_aba.strip()}') não cita {mes_chave}."))
    if str(ano) not in titulo:
        problemas.append(("AVISO", f"título da aba não cita o ano {ano}."))

    # 2) Proporções entre 0 e 1.
    for chave, valor in dados.items():
        if "__" in chave or not chave.startswith(PREFIXOS_PROPORCAO) or not _num(valor):
            continue
        if not 0 <= valor <= 1:
            problemas.append(("ERRO", f"{chave} = {valor} fora de 0–100% (fórmula da planilha?)."))

    # 3) Taxa = numerador / denominador.
    for taxa, num, den in RAZOES:
        if all(_num(dados.get(k)) for k in (taxa, num, den)) and dados[den]:
            esperado = dados[num] / dados[den]
            if abs(esperado - dados[taxa]) > 0.005 * max(1, abs(esperado)):
                problemas.append(("ERRO", f"{taxa} = {dados[taxa]:.4f}, mas {num}/{den} = {esperado:.4f}."))

    # 4) Consistência entre fontes.
    for descricao, a, b in CRUZADOS:
        if _num(dados.get(a)) and _num(dados.get(b)) and dados[a] != dados[b]:
            problemas.append(("AVISO", f"{descricao}: {dados[a]:g} x {dados[b]:g}."))

    # 5) TOTAL da planilha = soma dos dias (linhas de contagem). Divergência indica fórmula quebrada.
    for chave, serie in series.items():
        valores = [v for v in serie if v is not None]
        total = dados.get(chave)
        if not _num(total) or len(valores) < 2 or "taxa" in chave.split("__")[0] and "__" not in chave:
            continue
        if abs(sum(valores) - total) > 0.5:
            problemas.append(("ERRO", f"{chave}: TOTAL {total:g} ≠ soma dos dias {sum(valores):g} (fórmula da planilha?)."))

    # 6) Limitação do modelo: tabelas diárias do Word têm 30 colunas de dia.
    n_dias = calendar.monthrange(int(ano), list(MESES).index(mes_chave) + 1)[1]
    if n_dias > 30:
        problemas.append(("AVISO", f"{mes_chave} tem {n_dias} dias; as tabelas diárias do modelo têm 30 colunas."))
    return problemas
