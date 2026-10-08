"""Extração do Relatório Mensal da Fisioterapia (PDF): indicadores de ventilação mecânica.

A primeira tabela (Admissões / Fisioterapia Respiratória / Motora) é ignorada. Procura, em todas
as páginas, três seções pelo título e extrai o número associado a cada uma:
  - "Taxa de Ventilação Mecânica Invasiva total"        -> porcentagem (0–1)
  - "EXTUBAÇÃO NÃO PROGRAMADA"                          -> nº de ocorrências listadas
  - "PNEUMONIA ASSOCIADA À VENTILAÇÃO MECÂNICA – PAV"   -> nº de casos
"""

import re
from pathlib import Path

import pdfplumber

import log_execucao
from utils import sem_acento

log = log_execucao.obter(__name__)

CHAVES = ("fisio_taxa_vm_invasiva_total", "fisio_extubacao_nao_programada", "fisio_pav")
SEM_DADOS = "[Sem Dados]"

# (chave, título da seção normalizado). A seção vai do título ao próximo título da lista.
SECOES = (
    ("fisio_taxa_vm_invasiva_total", "TAXA DE VENTILACAO MECANICA INVASIVA TOTAL"),
    ("fisio_extubacao_nao_programada", "EXTUBACAO NAO PROGRAMADA"),
    ("fisio_pav", "PNEUMONIA ASSOCIADA A VENTILACAO MECANICA - PAV"),
)


def _normalizar(texto):
    """Normalização por linha do PDF: só colapsa espaços/tabs (utils.sem_acento já troca travessão por hífen)."""
    return re.sub(r"[ \t]+", " ", sem_acento(texto).upper())


def _numero(texto):
    return float(texto.replace(".", "").replace(",", "."))


def _taxa(secao):
    """Primeira porcentagem da seção, como fração: '71%' -> 0.71; '55,08%' -> 0.5508."""
    m = re.search(r"(\d{1,3}(?:[.,]\d+)?)\s*%", secao)
    return _numero(m.group(1)) / 100 if m else None


def _extubacoes(secao):
    """Ocorrências: itens '- dd/mm/aaaa ...' ou, sem itens, o número que abre a seção ('0 extubações')."""
    itens = re.findall(r"^\s*-\s*\d{2}/\d{2}/\d{4}", secao, flags=re.M)
    if itens:
        return len(itens)
    m = re.search(r"^\s*(\d+)\s+\S", secao, flags=re.M)
    return int(m.group(1)) if m else None


def _pav(secao):
    """'0 PAV caracterizada no mês...' -> 0."""
    m = re.search(r"(\d+)\s+PAV\b", secao)
    return int(m.group(1)) if m else None


_PARSERS = {
    "fisio_taxa_vm_invasiva_total": _taxa,
    "fisio_extubacao_nao_programada": _extubacoes,
    "fisio_pav": _pav,
}


def extrair_dados_fisio(caminho_pdf):
    """Devolve {fisio_taxa_vm_invasiva_total, fisio_extubacao_nao_programada, fisio_pav}.

    Indicador não localizado vira '[Sem Dados]' (com aviso), sem derrubar o restante.
    """
    caminho = Path(caminho_pdf)
    if not caminho.is_file():
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho}")

    try:
        with pdfplumber.open(caminho) as pdf:
            texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    except Exception as e:
        raise ValueError(f"Não foi possível ler '{caminho.name}': {e}") from e

    # Posição de cada título no texto (normalizado, mesmo comprimento do original por linha).
    linhas = texto.split("\n")
    norm = [_normalizar(l) for l in linhas]
    inicios = {}
    for chave, titulo in SECOES:
        for i, l in enumerate(norm):
            if l.strip().startswith(titulo):
                inicios[chave] = i
                break

    ordem = sorted(inicios.items(), key=lambda kv: kv[1])
    resultado = {}
    for chave, _ in SECOES:
        if chave not in inicios:
            log.warning(f"[WARNING] Fisioterapia: título do indicador '{chave}' não encontrado em '{caminho.name}'.")
            resultado[chave] = SEM_DADOS
            continue
        ini = inicios[chave] + 1
        fim = next((i for c, i in ordem if i > inicios[chave]), len(linhas))
        valor = _PARSERS[chave]("\n".join(linhas[ini:fim]))
        if valor is None:
            log.warning(f"[WARNING] Fisioterapia: valor de '{chave}' não localizado na seção do PDF.")
        resultado[chave] = SEM_DADOS if valor is None else valor
    return resultado


if __name__ == "__main__":
    import sys

    log_execucao.configurar()
    print(extrair_dados_fisio(sys.argv[1]))
