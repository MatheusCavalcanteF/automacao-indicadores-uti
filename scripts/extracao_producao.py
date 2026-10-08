"""Extração dos totais mensais da PLANILHA INDICADORES PRODUÇÃO DIÁRIA."""

import re
from pathlib import Path

import pandas as pd

from utils import MESES, normalizar as _normalizar, periodo_da_cli, resolver_mes, sem_acento

LINHA_CABECALHO = "DATAS"
COLUNA_TOTAL = "TOTAL"
COLUNA_REALIZADO = "REALIZADO"  # taxas/médias trazem o valor aqui; em TOTAL vem o texto "TOTAL"


def chave_variavel(texto):
    """'N. DE SAÍDAS' -> 'n_de_saidas' (variável Jinja válida). Chaves já normalizadas passam intactas."""
    if re.fullmatch(r"[a-z0-9_]+", str(texto)):
        return str(texto)
    return re.sub(r"[^a-z0-9]+", "_", sem_acento(texto).lower()).strip("_")


def _encontrar_aba(abas, mes_referencia):
    """Retorna o nome da aba que corresponde ao mês (nome completo ou abreviação)."""
    candidatos = set(MESES[resolver_mes(mes_referencia)][:2])

    # Prioriza o nome completo; só depois a abreviação (como palavra inteira).
    # Abas "PREENCHIMENTO GESTORA - SET" também casam: as "UTI PEDIATRICA" têm preferência.
    for termo in sorted(candidatos, key=len, reverse=True):
        achadas = [a for a in abas if re.search("(?<![A-Z])" + re.escape(termo) + "(?![A-Z])", _normalizar(a))]
        achadas.sort(key=lambda a: "UTI PEDIATRICA" not in _normalizar(a))
        if achadas:
            return achadas[0]
    raise ValueError(f"Nenhuma aba corresponde ao mês '{mes_referencia}'. Abas: {abas}")


def _localizar_cabecalho(bruto):
    """Acha (linha, coluna) da célula 'DATAS' no DataFrame lido sem cabeçalho."""
    for i, linha in bruto.iterrows():
        for j, valor in linha.items():
            if isinstance(valor, str) and _normalizar(valor) == LINHA_CABECALHO:
                return i, j
    raise ValueError(f"Linha '{LINHA_CABECALHO}' não encontrada na aba.")


def ler_producao(caminho_arquivo, mes_referencia):
    """Lê a aba do mês e devolve (totais, series, titulo_aba, meta).

    totais: {chave_variavel: valor} para toda linha com valor numérico. O valor vem de TOTAL ou,
        se não for numérico, de REALIZADO. Linhas com TOTAL não numérico são títulos de bloco e
        prefixam as linhas seguintes ('bloco__indicador'); linha em branco encerra o bloco.
        Chaves repetidas ganham sufixo _2, _3...
    series: {chave_variavel: [valor do dia 1, dia 2, ...]} (None onde a célula não é numérica).
    titulo_aba: texto das linhas acima de 'DATAS' (para conferir mês/ano da aba).
    meta: {"aba": nome da aba, "col_rotulo": coluna do rótulo (1 = A), "celulas": {chave: {linha, col_valor,
        col_ini, col_fim, rotulo}}} — onde cada indicador mora no Excel, para corrigir a planilha.
    """
    caminho = Path(caminho_arquivo)
    if not caminho.is_file():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    try:
        planilha = pd.ExcelFile(caminho, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Não foi possível abrir '{caminho.name}': {e}") from e

    with planilha:
        aba = _encontrar_aba(planilha.sheet_names, mes_referencia)
        bruto = planilha.parse(aba, header=None)

    # Cabeçalhos mesclados e linhas em branco: acha a linha 'DATAS' em vez de fixar o skiprows.
    lin_cab, col_rotulo = _localizar_cabecalho(bruto)
    titulo_aba = " ".join(str(v) for v in bruto.iloc[:lin_cab].to_numpy().ravel() if pd.notna(v))
    cabecalho = bruto.iloc[lin_cab].map(lambda v: _normalizar(v) if pd.notna(v) else "")
    colunas_total = [c for c, v in cabecalho.items() if v == COLUNA_TOTAL]
    if not colunas_total:
        raise ValueError(f"Coluna '{COLUNA_TOTAL}' não encontrada na aba '{aba}'.")
    col_total = colunas_total[0]
    cols_realizado = [c for c, v in cabecalho.items() if v == COLUNA_REALIZADO]
    col_realizado = cols_realizado[0] if cols_realizado else None
    cols_dias = [c for c in bruto.columns if col_rotulo < c < col_total]

    dados = bruto.iloc[lin_cab + 1:]
    totais, series, celulas = {}, {}, {}
    bloco = ""  # título do bloco pai; linha em branco encerra o bloco
    for idx, rotulo in dados[col_rotulo].items():
        nome = _normalizar(rotulo) if pd.notna(rotulo) else ""
        if not nome:
            bloco = ""
            continue

        total = pd.to_numeric(dados.at[idx, col_total], errors="coerce")
        realizado = (
            pd.to_numeric(dados.at[idx, col_realizado], errors="coerce") if col_realizado is not None else float("nan")
        )
        valor = total if pd.notna(total) else realizado

        # TOTAL não numérico: a linha é título de bloco (o valor, se houver, vem em REALIZADO).
        if pd.isna(total):
            bloco = nome
            chave = chave_variavel(nome)
        else:
            chave = chave_variavel(bloco) + "__" + chave_variavel(nome) if bloco else chave_variavel(nome)

        if pd.isna(valor) or not chave:
            continue
        base, n = chave, 2
        while chave in totais:  # colisão residual: sufixo incremental
            chave = f"{base}_{n}"
            n += 1
        totais[chave] = float(valor)
        dias = [pd.to_numeric(dados.at[idx, c], errors="coerce") for c in cols_dias]
        series[chave] = [None if pd.isna(v) else float(v) for v in dias]
        # Índices do pandas são 0-based; no Excel linhas e colunas começam em 1.
        celulas[chave] = {
            "linha": int(idx) + 1,
            "col_valor": int(col_total if pd.notna(total) else col_realizado) + 1,
            "col_ini": int(cols_dias[0]) + 1,
            "col_fim": int(cols_dias[-1]) + 1,
            "rotulo": str(rotulo),
        }
    meta = {"aba": aba, "col_rotulo": int(col_rotulo) + 1, "celulas": celulas}
    return totais, series, titulo_aba, meta


def extrair_dados_diarios(caminho_arquivo, mes_referencia):
    """Só os totais: {chave_variavel: valor} (ver ler_producao)."""
    return ler_producao(caminho_arquivo, mes_referencia)[0]


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 3:
        sys.exit("Uso: python scripts/extracao_producao.py PLANILHA.xlsx MES")
    periodo = periodo_da_cli(sys.argv[2:3], "python scripts/extracao_producao.py PLANILHA.xlsx MES")
    totais, series, titulo, _ = ler_producao(sys.argv[1], periodo.mes)
    print(f"Foram extraídos {len(totais)} indicadores; título da aba: {titulo!r}")
