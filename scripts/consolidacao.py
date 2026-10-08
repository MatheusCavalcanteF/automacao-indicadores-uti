"""Localiza as fontes em dados_entrada/, extrai, concilia/sanitiza e consolida os dicionários."""

import hashlib
from datetime import datetime
from pathlib import Path

import log_execucao
from extracao_censo import CHAVES as CHAVES_CENSO, extrair_dados_censo
from extracao_odt_testes import CHAVES as CHAVES_TESTES, extrair_dados_testes
from extracao_pdf_fisio import CHAVES as CHAVES_FISIO, extrair_dados_fisio
from extracao_producao import ler_producao
from localizar_arquivos import ArquivoAmbiguoError, localizar_arquivo
from sanitizacao import sanitizar

RAIZ = Path(__file__).resolve().parent.parent
ENTRADA = RAIZ / "dados_entrada"
PENDENTE = "[Sem Dados]"

log = log_execucao.obter(__name__)

# (título da seção, extensão, termos no nome, obrigatória?, variáveis para marcar como pendentes)
FONTES = [
    ("PRODUÇÃO DIÁRIA", ".xlsx", ["produ"], True, ()),
    ("CENSO", ".xlsx", ["censo"], False, CHAVES_CENSO),
    ("FISIOTERAPIA", ".pdf", ["fisio"], False, CHAVES_FISIO),
    ("TESTES (ORELHINHA / LINGUINHA)", ".odt", ["nascidos", "testes"], False, CHAVES_TESTES),
]


def _sha256(caminho):
    """Impressão digital do arquivo de entrada, para a auditoria provar qual versão foi usada."""
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(1 << 20), b""):
            h.update(bloco)
    return h.hexdigest()


def _mtime(caminho):
    return f"{datetime.fromtimestamp(caminho.stat().st_mtime):%Y-%m-%d %H:%M:%S}"


def extrair_por_fonte(mes, ano, pasta_entrada=ENTRADA):
    """Retorna ({seção: {variável: valor}}, auxiliar) e imprime o arquivo usado em cada fonte.

    Só a Produção é obrigatória. Se uma fonte secundária faltar, ou só existir de outro mês,
    emite [WARNING] e preenche as variáveis dela com "[Sem Dados]" (nunca usa dado de outro mês).

    auxiliar: {"series": séries diárias da produção, "titulo_aba": texto acima de DATAS,
               "arquivos": {seção: nome do arquivo}, "pendencias": [seções sem dados],
               "meta": onde cada indicador está no Excel, "arquivo_producao": Path da planilha}.
    """
    secoes, arquivos, pendencias = {}, {}, []
    auxiliar = {"series": {}, "titulo_aba": "", "arquivos": arquivos, "pendencias": pendencias,
                 "meta": None, "arquivo_producao": None}

    for titulo, extensao, termos, obrigatoria, chaves in FONTES:
        try:
            arquivo = localizar_arquivo(
                pasta_entrada, extensao, termos, mes, ano, descricao=titulo, estrito=not obrigatoria
            )
            log.info(f"  {titulo}: {arquivo.name}")
            log.debug(f"  {titulo}: sha256 {_sha256(arquivo)}; modificado em {_mtime(arquivo)}")
            arquivos[titulo] = arquivo.name
            if titulo == "PRODUÇÃO DIÁRIA":
                secoes[titulo], auxiliar["series"], auxiliar["titulo_aba"], auxiliar["meta"] = ler_producao(arquivo, mes)
                auxiliar["arquivo_producao"] = arquivo
            elif titulo == "CENSO":
                secoes[titulo] = extrair_dados_censo(arquivo)
            elif titulo == "FISIOTERAPIA":
                secoes[titulo] = extrair_dados_fisio(arquivo)
            else:
                secoes[titulo] = extrair_dados_testes(arquivo)
        except (FileNotFoundError, ArquivoAmbiguoError, ValueError) as e:
            if obrigatoria:
                raise
            motivo = str(e) if titulo in str(e) else f"{titulo}: {e}"
            log.warning(f"[WARNING] {motivo}\n          Variáveis desta fonte ficam como '{PENDENTE}'.")
            arquivos.pop(titulo, None)
            pendencias.append(titulo)
            secoes[titulo] = {chave: PENDENTE for chave in chaves}
    return secoes, auxiliar


def consolidar(secoes):
    """Junta as seções em um único dicionário, avisando sobre chaves repetidas."""
    dados_consolidados = {}
    for titulo, dados in secoes.items():
        repetidas = dados_consolidados.keys() & dados.keys()
        if repetidas:
            log.warning(f"[aviso] '{titulo}' sobrescreve chaves já existentes: {sorted(repetidas)}")
        dados_consolidados.update(dados)
    return dados_consolidados


def consolidar_e_sanitizar(secoes, series=None):
    """consolidar() + sanitização (Produção como fonte da verdade, fórmulas recalculadas).

    Retorna (dados, log, correcoes): log = [(nível, mensagem)]; correcoes = ações estruturadas
    que corrigir_planilha() grava no Excel.
    """
    return sanitizar(consolidar(secoes), series)
