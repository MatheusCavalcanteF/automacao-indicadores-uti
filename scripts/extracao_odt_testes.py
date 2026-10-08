"""Extração da tabela de testes (Orelhinha/Linguinha) do ODT de Nascidos Vivos. Ignora texto livre."""

from pathlib import Path

from odf import teletype
from odf.opendocument import load
from odf.table import Table, TableCell, TableRow

from extracao_producao import chave_variavel
from utils import normalizar as _normalizar

LINHA_ALVO = "MATERNIDADE"  # padrão histórico dos relatórios
COLUNAS = {"teste_orelhinha": "TESTE ORELHINHA", "teste_linguinha": "TESTE LINGUINHA"}
SETORES = ("uti_ped", "uti_neo", "pediatria", "maternidade", "ambulatorio", "total")
# Todas as variáveis que o modelo pode usar (para preencher como pendente se o arquivo faltar).
CHAVES = tuple(COLUNAS) + tuple(f"{c}_{s}" for c in COLUNAS for s in SETORES)


def _texto(no):
    return " ".join(teletype.extractText(no).split())


def _linhas_da_tabela(tabela):
    linhas = []
    for linha in tabela.getElementsByType(TableRow):
        celulas = []
        for celula in linha.getElementsByType(TableCell):
            repete = int(celula.getAttribute("numbercolumnsrepeated") or 1)
            celulas.extend([_texto(celula)] * min(repete, 50))
        if any(celulas):
            linhas.append(celulas)
    return linhas


def extrair_dados_testes(caminho_odt):
    """Devolve {teste_orelhinha, teste_linguinha} da linha MATERNIDADE, mais uma variável por setor
    (teste_orelhinha_uti_neo, teste_linguinha_total, ...) para a tabela completa do relatório."""
    caminho = Path(caminho_odt)
    if not caminho.is_file():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    try:
        doc = load(str(caminho))
    except Exception as e:
        raise ValueError(f"Não foi possível ler '{caminho.name}': {e}") from e

    for tabela in doc.getElementsByType(Table):
        linhas = _linhas_da_tabela(tabela)
        cab = next((l for l in linhas if any("ORELHINHA" in _normalizar(c) for c in l)), None)
        if cab is None:
            continue
        cab = [_normalizar(c) for c in cab]
        for rotulo in COLUNAS.values():
            if rotulo not in cab:
                raise ValueError(f"Coluna '{rotulo}' não encontrada. Cabeçalho: {cab}")

        def numero(bruto, onde):
            try:
                return int(float(bruto.replace(".", "").replace(",", ".")))
            except ValueError:
                raise ValueError(f"Valor não numérico em {onde}: {bruto!r}") from None

        resultado = {}
        for linha in linhas:
            setor = _normalizar(linha[0])
            if setor == _normalizar(cab[0]) or len(linha) < len(cab):
                continue  # cabeçalho ou linha incompleta
            for chave, rotulo in COLUNAS.items():
                valor = numero(linha[cab.index(rotulo)], f"{chave} ({setor})")
                resultado[f"{chave}_{chave_variavel(setor)}"] = valor
                if setor == LINHA_ALVO:
                    resultado[chave] = valor
        if "teste_orelhinha" in resultado:
            return resultado
    raise ValueError(f"Linha '{LINHA_ALVO}' não encontrada nas tabelas de '{caminho.name}'.")


if __name__ == "__main__":
    import sys

    print(extrair_dados_testes(sys.argv[1]))
