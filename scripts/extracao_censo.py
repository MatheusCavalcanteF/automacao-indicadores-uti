"""Extração do resumo do Censo (09 CENSO INPUT) para a linha UTI Pediátrica."""

from pathlib import Path

import pandas as pd

from utils import normalizar as _normalizar

ABA = "Resumo"
LINHAS_CABECALHO = (1, 2, 3)  # cabeçalho mesclado: rótulo se espalha por estas linhas
LINHA_ALVO = "UTI PEDIATRICA"  # a especialidade deve conter este texto

# nome no projeto -> texto do cabeçalho (normalizado, linhas do cabeçalho unidas)
COLUNAS = {
    "censo_leitos_dia": "LEITOS DIA",
    "censo_paciente_dia": "PACIENTE DIA",
    "censo_total_saidas_geral": "TOTAL SAIDAS GERAL",
}

CHAVES = tuple(COLUNAS)


def extrair_dados_censo(caminho_arquivo):
    """Lê a aba 'Resumo' e devolve os valores da linha UTI Pediátrica."""
    caminho = Path(caminho_arquivo)
    if not caminho.is_file():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    try:
        bruto = pd.read_excel(caminho, sheet_name=ABA, header=None, engine="openpyxl")
    except ValueError as e:
        raise ValueError(f"Aba '{ABA}' não encontrada em '{caminho.name}': {e}") from e

    # Junta as linhas do cabeçalho mesclado: "Leitos" + "Dia" -> "LEITOS DIA".
    rotulos = {
        col: _normalizar(" ".join(str(bruto.at[l, col]) for l in LINHAS_CABECALHO if pd.notna(bruto.at[l, col])))
        for col in bruto.columns
    }

    # Coluna de especialidade: a que contém a linha-alvo.
    linha = col_esp = None
    for col in bruto.columns:
        achadas = bruto.index[bruto[col].map(lambda v: isinstance(v, str) and LINHA_ALVO in _normalizar(v))]
        if len(achadas):
            linha, col_esp = achadas[0], col
            break
    if linha is None:
        raise ValueError(f"Linha '{LINHA_ALVO}' não encontrada na aba '{ABA}'.")

    resultado = {}
    for nome, rotulo in COLUNAS.items():
        cols = [c for c, r in rotulos.items() if r == rotulo]
        if not cols:
            raise ValueError(f"Coluna '{nome}' não encontrada no cabeçalho. Cabeçalhos: {sorted(set(rotulos.values()))}")
        valor = pd.to_numeric(bruto.at[linha, cols[0]], errors="coerce")
        resultado[nome] = None if pd.isna(valor) else float(valor)
    return resultado


if __name__ == "__main__":
    import sys

    print(extrair_dados_censo(sys.argv[1]))
