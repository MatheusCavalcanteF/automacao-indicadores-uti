"""Gráficos anuais do relatório: valores das abas mensais da Produção, desenho e troca da imagem no Word.

Fonte: cada aba mensal, de janeiro até o mês de referência, com a mesma sanitização do relatório. O mês
de referência usa os próprios dados do relatório, para a barra bater com a tabela. A aba FECHAMENTO
COMPILADO ANUAL não é usada: tem #REF! e colunas apontando para o mês errado.

O modelo é ligado por marcador, não por posição: montar_template grava no texto alternativo da imagem
"grafico:<chave>;meta=<valor>", lendo a chave e a meta (célula "Contratado") da tabela do indicador.
As imagens do documento-base estão em VML (v:imagedata), que o replace_pic do docxtpl não enxerga; por
isso a troca é feita aqui, nas relações da parte principal.

Mês sem dado válido sai como "s/d", nunca como 0. Meses depois do de referência ficam vazios.
"""

import io
import re
import textwrap
import zipfile
from collections import namedtuple
from pathlib import Path

from docx.oxml.ns import qn
from matplotlib import font_manager
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
from matplotlib.ticker import FixedLocator, FuncFormatter, MaxNLocator

import log_execucao
from extracao_producao import ler_meses
from sanitizacao import sanitizar
from utils import MESES, eh_numero

log = log_execucao.obter(__name__)

MESES_ROTULO = ("JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO",
                "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO")
SEM_DADO = "s/d"
# As cores dos gráficos do relatório, para o documento não mudar de aparência.
COR_BARRA, COR_META, COR_TEXTO, COR_GRADE, COR_SEM_DADO = "#FFC000", "#FF0000", "#404040", "#D9D9D9", "#7F7F7F"
TAMANHO_PADRAO = (752, 434)  # pixels da imagem original, usados se ela não puder ser lida
DPI = 200  # o dobro da resolução original, para não pixelar na impressão

Grafico = namedtuple("Grafico", "titulo formato casas")
Grafico.__doc__ = "formato: 'pct' (fração exibida em %) ou 'dec'; casas: decimais do rótulo de cada barra."

# Mesmas chaves de montar_template.TABELAS.
GRAFICOS = {
    "saidas_hospitalares__n_de_saidas": Grafico("SAÍDAS HOSPITALARES", "dec", 0),
    "taxa_de_ocupacao_sem_os_cronicos": Grafico("TAXA DE OCUPAÇÃO SEM OS PACIENTES CRÔNICOS", "pct", 0),
    "taxa_de_ocupacao_global": Grafico("TAXA DE OCUPAÇÃO GLOBAL", "pct", 0),
    "media_de_permanencia_sem_os_cronicos": Grafico("MÉDIA DE PERMANÊNCIA - SEM OS PACIENTES CRÔNICOS", "dec", 1),
    "media_de_permanencia": Grafico("MÉDIA DE PERMANÊNCIA - GLOBAL", "dec", 1),
    "taxa_de_mortalidade_em_uti": Grafico("TAXA DE MORTALIDADE", "pct", 2),
    "atender_100_das_solicitacoes_de_interconsultas_em_neuro": Grafico(
        "ATENDER 100% DAS SOLICITAÇÕES DE INTERCONSULTAS EM NEUROLOGIA", "pct", 0),
    "taxa_de_reinternacao_em_24h": Grafico("TAXA DE REINTERNAÇÃO EM 24 HORAS", "pct", 2),
    "pnm_associada_a_vm_pav": Grafico("% PNM ASSOCIADA A VM (PAV) ≤", "pct", 2),
    "taxa_de_utilizacao_de_ventilacao_mecanica": Grafico("TAXA DE UTILIZAÇÃO DE VENTILAÇÃO MECÂNICA", "pct", 0),
    "incidencia_de_extubacao_acidental": Grafico("INCIDÊNCIA DE EXTUBAÇÃO ACIDENTAL", "pct", 2),
    "incidencia_de_saida_nao_planejada_de_sonda_oro_nasogastroenteral": Grafico(
        "INCIDÊNCIA DE SAÍDA NÃO PLANEJADA DE SONDA ORO/NASOGASTROENTERAL ≤", "pct", 2),
    "densidade_de_incidencia_ipcs": Grafico("DENSIDADE DE INCIDÊNCIA IPCS ≤", "dec", 2),
    "incidencia_de_flebite": Grafico("INCIDÊNCIA DE FLEBITE ≤", "pct", 2),
    "incidencia_de_perda_de_cateter_central_insercao_periferica": Grafico(
        "INCIDÊNCIA DE PERDA DE CATETER CENTRAL INSERÇÃO PERIFÉRICA ≤", "pct", 2),
    "incidencia_de_perda_de_cateter_central": Grafico("INCIDÊNCIA DE PERDA DE CATETER CENTRAL ≤", "pct", 2),
    "taxa_de_utilizacao_de_cateter_venoso_central_cvc": Grafico(
        "TAXA DE UTILIZAÇÃO DE CATETER VENOSO CENTRAL (CVC) ≤", "pct", 0),
    "incidencia_de_erro_quase_erro_de_medicacao": Grafico("INCIDÊNCIA DE ERRO/QUASE ERRO DE MEDICAÇÃO ≤", "pct", 2),
    "densidade_de_incidencia_itu": Grafico("DENSIDADE DE INCIDÊNCIA ITU ≤", "dec", 2),
    "incidencia_de_queda_de_paciente": Grafico("INCIDÊNCIA DE QUEDA DE PACIENTE ≤", "pct", 2),
    "indice_de_ulcera_por_pressao": Grafico("ÍNDICE DE ÚLCERA POR PRESSÃO <", "pct", 2),
    "adesao_as_metas_de_identificacao_do_paciente": Grafico("ADESÃO ÀS METAS DE IDENTIFICAÇÃO DO PACIENTE", "pct", 0),
    "taxa_de_adesao_ao_protocolo_de_hm_momentos": Grafico("TAXA DE ADESÃO AO PROTOCOLO DE HM - 05 MOMENTOS", "pct", 0),
    "evento_sentinela_0": Grafico("EVENTO SENTINELA", "dec", 0),
    "atender_100_das_solicitacoes_de_transporte_inter_e_intrahospitalar": Grafico(
        "ATENDER 100% DAS SOLICITAÇÕES DE TRANSPORTE INTER E INTRAHOSPITALAR", "pct", 0),
    "fonoaudiologo_teste_triagem_auditiva_universal_teste_da_orelhinha_100_dos_recem_nascidos": Grafico(
        "FONOAUDIÓLOGO TESTE TRIAGEM AUDITIVA UNIVERSAL (TESTE DA ORELHINHA) / 100% DOS RECÉM-NASCIDOS", "pct", 0),
    "fonoaudiologo_teste_linguinha_em_100_dos_nascidos": Grafico(
        "FONOAUDIÓLOGO TESTE LINGUINHA EM 100% DOS RECÉM-NASCIDOS", "pct", 0),
}
# A célula "Contratado" da neurologia traz um prazo, não a meta do gráfico.
META_FORA_DO_CONTRATADO = {"atender_100_das_solicitacoes_de_interconsultas_em_neuro": 1.0}

HistoricoAnual = namedtuple("HistoricoAnual", "ano n_meses series")
HistoricoAnual.__doc__ = "series: {chave do gráfico: 12 valores}; None = sem dado válido ou mês futuro."

_NS_VML = "{urn:schemas-microsoft-com:vml}"
_PREFIXO_MARCADOR = "grafico:"
_MARCADOR = re.compile(r"^grafico:(?P<chave>[a-z0-9_]+)(?:;meta=(?P<meta>[0-9.eE+-]+))?$")
_NUMERO_META = re.compile(r"\d+(?:,\d+)?")


# ---------------------------------------------------------------- valores

def historico_anual(caminho_producao, periodo, dados):
    """Monta o HistoricoAnual de janeiro até `periodo`; o mês de referência vem de `dados` já sanitizados.

    Aba ausente ou ilegível deixa o mês sem dado. Ajustes da sanitização nos outros meses vão só ao log
    de auditoria: os avisos do relatório tratam do mês de referência.
    """
    anteriores = list(MESES)[:periodo.numero - 1]
    lidos = ler_meses(caminho_producao, anteriores) if anteriores else {}
    por_mes = []
    for mes in anteriores:
        if mes not in lidos:
            log.debug(f"[GRÁFICO] {mes}: aba não encontrada ou ilegível; mês sem dados nos gráficos.")
            por_mes.append({})
            continue
        valores, registro, _ = sanitizar(*lidos[mes])
        for nivel, msg in registro:
            log.debug(f"[GRÁFICO] {mes} [{nivel}] {msg}")
        por_mes.append(valores)
    por_mes.append(dados)

    series = {}
    for chave in GRAFICOS:
        valores = [v if eh_numero(v := mes.get(chave)) else None for mes in por_mes]
        series[chave] = valores + [None] * (12 - len(valores))
    return HistoricoAnual(periodo.ano, periodo.numero, series)


def meta_do_contratado(texto, filtro, chave=None):
    """'>=80%' -> 0.8; '<=6,5 dias' -> 6.5; '< 0,25' (filtro pct) -> 0.0025. Sem número, devolve None."""
    if chave in META_FORA_DO_CONTRATADO:
        return META_FORA_DO_CONTRATADO[chave]
    achado = _NUMERO_META.search(texto or "")
    if not achado:
        return None
    valor = float(achado.group().replace(",", "."))
    return round(valor / 100, 6) if filtro == "pct" else valor


# ---------------------------------------------------------------- desenho

def _br(valor, casas):
    return f"{valor:.{casas}f}".replace(".", ",")


def rotulo_valor(valor, grafico):
    """Rótulo da barra. Valor inteiro sai sem casas ('0%', não '0,00%')."""
    numero = valor * 100 if grafico.formato == "pct" else valor
    casas = 0 if abs(numero - round(numero)) < 1e-9 else grafico.casas
    return _br(numero, casas) + ("%" if grafico.formato == "pct" else "")


def rotulo_meta(valor, grafico):
    """Rótulo da meta, com as casas de que ela precisa: 80%, 41,27%, 6,5."""
    numero = valor * 100 if grafico.formato == "pct" else valor
    texto = f"{round(numero, 2):.2f}".rstrip("0").rstrip(".").replace(".", ",")
    return texto + ("%" if grafico.formato == "pct" else "")


def _fonte():
    disponiveis = {f.name for f in font_manager.fontManager.ttflist}
    return next((f for f in ("Calibri", "Carlito") if f in disponiveis), "DejaVu Sans")


def _formatar_eixo(eixo, grafico):
    """Eixo Y a partir de zero, com os decimais que os ticks pedem e a mesma unidade dos rótulos."""
    escala = 100 if grafico.formato == "pct" else 1
    ticks = [t for t in eixo.get_yticks() if t >= 0 and t <= eixo.get_ylim()[1]]
    casas = next((c for c in range(4) if all(abs(round(t * escala, c) - t * escala) < 1e-9 for t in ticks)), 2)
    sufixo = "%" if grafico.formato == "pct" else ""
    eixo.yaxis.set_major_locator(FixedLocator(ticks))
    eixo.yaxis.set_major_formatter(FuncFormatter(lambda v, _: _br(v * escala, casas) + sufixo))


def figura_grafico(grafico, valores, n_meses, ano, meta=None, tamanho=TAMANHO_PADRAO):
    """Figura do gráfico: barras de janeiro a dezembro e linha de meta. n_meses=0 = dados indisponíveis."""
    figura = Figure(figsize=(tamanho[0] / 100, tamanho[1] / 100), dpi=DPI, facecolor="white")
    FigureCanvasAgg(figura)
    fonte = _fonte()
    eixo = figura.add_subplot()

    numeros = [v for v in valores if v is not None]
    posicoes = [i for i, v in enumerate(valores) if v is not None]
    referencia = numeros + ([meta] if meta is not None else [])
    topo = max(referencia, default=0) * 1.18 or 1
    eixo.set_ylim(0, topo)
    eixo.set_xlim(-0.6, 12.0 if meta is not None else 11.6)
    eixo.bar(posicoes, numeros, width=0.42, color=COR_BARRA, zorder=2)
    if meta is not None:
        eixo.plot([0, 11], [meta, meta], color=COR_META, linewidth=2, zorder=3)
        eixo.text(11.12, meta, rotulo_meta(meta, grafico), color=COR_META, fontsize=9, fontweight="bold",
                  va="center", ha="left", fontfamily=fonte, zorder=5)

    folga = topo * 0.012
    for i, valor in zip(posicoes, numeros):
        eixo.text(i, valor + folga, rotulo_valor(valor, grafico), ha="center", va="bottom", fontsize=8.5,
                  fontweight="bold", color=COR_TEXTO, fontfamily=fonte, zorder=4,
                  bbox={"facecolor": "white", "alpha": 0.7, "edgecolor": "none", "pad": 0.4})
    faltam = [i for i in range(n_meses) if valores[i] is None]
    for i in faltam:
        eixo.text(i, folga, SEM_DADO, ha="center", va="bottom", fontsize=8.5, color=COR_SEM_DADO,
                  fontfamily=fonte, zorder=4)
    if n_meses == 0:
        eixo.text(0.5, 0.5, "Gráfico indisponível", transform=eixo.transAxes, ha="center", va="center",
                  fontsize=12, color=COR_SEM_DADO, fontfamily=fonte)

    eixo.set_xticks(range(12), MESES_ROTULO, rotation=45, ha="right", rotation_mode="anchor",
                    fontsize=7.5, color=COR_TEXTO, fontfamily=fonte)
    eixo.yaxis.set_major_locator(MaxNLocator(nbins=7, steps=[1, 2, 2.5, 5, 10]))
    _formatar_eixo(eixo, grafico)
    eixo.tick_params(axis="y", labelsize=8.5, labelcolor=COR_TEXTO, length=0)
    eixo.tick_params(axis="x", length=0)
    for rotulo in eixo.get_yticklabels():
        rotulo.set_fontfamily(fonte)
    eixo.grid(axis="y", color=COR_GRADE, linewidth=0.7, zorder=0)
    eixo.set_axisbelow(True)
    for lado in ("top", "right", "left"):
        eixo.spines[lado].set_visible(False)
    eixo.spines["bottom"].set_color(COR_GRADE)

    eixo.set_title("\n".join(textwrap.wrap(grafico.titulo, 58)), fontsize=12.5, fontweight="bold",
                   color=COR_TEXTO, fontfamily=fonte, pad=8)
    # Itens explícitos: mês sem nenhuma barra não pode deixar a legenda sem a cor da barra.
    if n_meses:
        itens = [Patch(color=COR_BARRA, label=str(ano))]
        if meta is not None:
            itens.append(Line2D([], [], color=COR_META, linewidth=2, label="META"))
        figura.legend(handles=itens, loc="lower center", ncol=2, frameon=False,
                      prop={"family": fonte, "size": 9}, labelcolor=COR_TEXTO)
    if faltam:  # não usar 7 pt: a 200 dpi, a Calibri nesse tamanho vira um borrão no canto da imagem
        figura.text(0.985, 0.02, f"{SEM_DADO} = sem dado válido na planilha", ha="right", va="bottom",
                    fontsize=7.5, color=COR_SEM_DADO, fontfamily=fonte)
    figura.subplots_adjust(left=0.075, right=0.985, top=0.86 if "\n" not in eixo.get_title() else 0.8, bottom=0.25)
    return figura


def png_do_grafico(*args, **kwargs):
    """PNG (bytes) de figura_grafico()."""
    saida = io.BytesIO()
    figura_grafico(*args, **kwargs).savefig(saida, format="png", metadata={"Software": None})
    return saida.getvalue()


# ---------------------------------------------------------------- Word

def _imagens(raiz):
    """(nó do texto alternativo, atributo, nó da imagem, atributo da relação) de cada imagem sob `raiz`."""
    for forma in raiz.iter(f"{_NS_VML}shape"):
        dados = forma.find(f"{_NS_VML}imagedata")
        if dados is not None and dados.get(qn("r:id")):
            yield forma, "alt", dados, qn("r:id")
    for marcador in (qn("wp:inline"), qn("wp:anchor")):
        for desenho in raiz.iter(marcador):
            blip = next(desenho.iter(qn("a:blip")), None)
            if blip is not None and blip.get(qn("r:embed")):
                yield desenho.find(qn("wp:docPr")), "descr", blip, qn("r:embed")


def marcar_imagem(elemento, chave, meta):
    """Grava o marcador na imagem de `elemento` (parágrafo). Devolve False se não houver imagem."""
    imagem = next(_imagens(elemento), None)
    if imagem is None:
        return False
    no_alt, atributo, _, _ = imagem
    no_alt.set(atributo, f"{_PREFIXO_MARCADOR}{chave}" + (f";meta={meta:.10g}" if meta is not None else ""))
    return True


def tem_imagem(elemento):
    return next(_imagens(elemento), None) is not None


def _tamanho_em_pixels(parte, rid):
    try:
        imagem = parte.related_parts[rid].image
        return imagem.px_width, imagem.px_height
    except Exception:  # noqa: BLE001 - formato que o python-docx não lê: o gráfico sai no tamanho padrão
        return TAMANHO_PADRAO


def substituir_graficos(doc, historico):
    """Troca as imagens marcadas de `doc` pelos gráficos do mês e descarta as imagens antigas do pacote.

    historico None (falha ao ler as abas): todas viram "Gráfico indisponível", nunca o gráfico antigo.
    Retorna a lista de mensagens de conferência sobre meses sem dado.
    """
    parte = doc.part
    n_meses = historico.n_meses if historico else 0
    ano = historico.ano if historico else ""
    series = historico.series if historico else {}
    avisos, antigas, trocadas = [], set(), 0

    for no_alt, atributo, no_imagem, relacao in list(_imagens(doc.element.body)):
        achado = _MARCADOR.match((no_alt.get(atributo) or "") if no_alt is not None else "")
        if not achado or achado["chave"] not in GRAFICOS:
            continue
        chave, grafico = achado["chave"], GRAFICOS[achado["chave"]]
        meta = float(achado["meta"]) if achado["meta"] else None
        valores = series.get(chave) or [None] * 12
        antiga = no_imagem.get(relacao)
        png = png_do_grafico(grafico, valores, n_meses, ano, meta, _tamanho_em_pixels(parte, antiga))
        nova, _ = parte.get_or_add_image(io.BytesIO(png))
        no_imagem.set(relacao, nova)
        no_alt.set(atributo, f"Gráfico: {grafico.titulo}")
        antigas.add(antiga)
        trocadas += 1
        faltam = [MESES_ROTULO[i].capitalize() for i in range(n_meses) if valores[i] is None]
        if faltam:
            avisos.append(f"Gráfico '{grafico.titulo}': sem dado válido na planilha em {', '.join(faltam)}; "
                          f"o mês aparece como '{SEM_DADO}'.")

    em_uso = set(parte.element.xpath("//@r:id | //@r:embed"))
    for rid in antigas - em_uso:
        parte.rels.pop(rid)
    if trocadas:
        log.info(f"[GRÁFICO] {trocadas} gráfico(s) gerado(s) a partir das abas mensais.")
    return avisos


def aviso_modelo_sem_marcadores(caminho_modelo):
    """Mensagem se o modelo tem imagens sem nenhum marcador (feito antes dos gráficos dinâmicos), senão None."""
    try:
        with zipfile.ZipFile(caminho_modelo) as z:
            xml = z.read("word/document.xml").decode("utf-8")
    except (OSError, KeyError, zipfile.BadZipFile):
        return None
    if f'"{_PREFIXO_MARCADOR}' in xml or not re.search(r"<v:imagedata\b|<a:blip\b", xml):
        return None
    return (f"Modelo {Path(caminho_modelo).name} sem marcadores de gráfico: as imagens são as do documento-base, "
            "de outro mês. Rode scripts/montar_template.py e gere de novo.")
