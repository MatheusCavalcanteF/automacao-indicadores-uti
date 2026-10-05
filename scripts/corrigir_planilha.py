"""Write-back: corrige as fórmulas quebradas na planilha de Produção e salva `<nome>_Corrigida.xlsx`.

Leitura e conferência com openpyxl; a escrita é um patch direto no XML da célula. Motivo: um
load/save do openpyxl reescreve o arquivo inteiro e descarta formas/caixas de texto e parte da
formatação dos gráficos (a planilha real tem 20 desenhos e 28 gráficos). O patch troca só a
fórmula (e o valor em cache) das células corrigidas; o resto do arquivo fica idêntico.
O original nunca é alterado.
"""

import re
import zipfile
from pathlib import Path

import openpyxl
from lxml import etree
from openpyxl.utils import get_column_letter

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
SUFIXO = "_Corrigida"


def _addr(celulas, chave):
    c = celulas[chave]
    return f"{get_column_letter(c['col_valor'])}{c['linha']}"


def _formula(acao, chave, celulas):
    c = celulas[chave]
    if acao[0] == "soma_linha":
        return f"=SUM({get_column_letter(c['col_ini'])}{c['linha']}:{get_column_letter(c['col_fim'])}{c['linha']})"
    if acao[0] == "div":
        return f"={_addr(celulas, acao[1])}/{_addr(celulas, acao[2])}"
    if acao[0] == "sub":
        return f"={_addr(celulas, acao[1])}-{_addr(celulas, acao[2])}"
    if acao[0] == "ref":
        return f"={_addr(celulas, acao[1])}"
    raise ValueError(f"Ação de correção desconhecida: {acao}")


def _canonica(formula):
    return re.sub(r"[\s$()]", "", str(formula)).upper()


def _planejar(caminho, aba, col_rotulo, celulas, correcoes):
    """Confere cada célula-alvo (rótulo na coluna A etc.) e monta as edições: addr -> dict."""
    wb = openpyxl.load_workbook(caminho)  # só leitura: nunca salvamos este objeto
    ws = wb[aba]
    edicoes = {}
    for corr in correcoes:
        chave = corr["chave"]
        if chave not in celulas:
            raise ValueError(f"Indicador '{chave}' sem célula rastreada.")
        info = celulas[chave]
        rotulo = ws.cell(info["linha"], col_rotulo).value
        if str(rotulo or "").strip() != info["rotulo"].strip():
            raise ValueError(f"{aba}!linha {info['linha']}: rótulo '{rotulo}' ≠ esperado '{info['rotulo']}' "
                             f"(planilha mudou desde a leitura?).")
        endereco = _addr(celulas, chave)
        atual = ws[endereco].value
        nova = _formula(corr["acao"], chave, celulas)
        mesma = isinstance(atual, str) and _canonica(atual) == _canonica(nova)
        edicoes[endereco] = {
            "chave": chave, "atual": atual, "formula": nova, "valor": float(corr["novo"]), "mesma": mesma,
        }
    return edicoes


def _caminho_da_aba(zf, aba):
    wb = etree.fromstring(zf.read("xl/workbook.xml"))
    rid = next((s.get(f"{{{NS_REL}}}id") for s in wb.iter(f"{{{NS}}}sheet") if s.get("name") == aba), None)
    if rid is None:
        raise ValueError(f"Aba '{aba}' não encontrada no workbook.xml.")
    rels = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    alvo = next(r.get("Target") for r in rels if r.get("Id") == rid)
    return alvo.lstrip("/") if alvo.startswith("/") else "xl/" + alvo


def _patch_aba(xml, edicoes):
    raiz = etree.fromstring(xml)
    celulas = {c.get("r"): c for c in raiz.iter(f"{{{NS}}}c")}
    for endereco, e in edicoes.items():
        c = celulas.get(endereco)
        if c is None:
            raise ValueError(f"Célula {endereco} não existe no XML da aba.")
        f = c.find(f"{{{NS}}}f")
        if f is not None and f.get("t") == "shared" and f.get("ref"):
            raise ValueError(f"{endereco} é a célula-mestre de uma fórmula compartilhada; correção manual necessária.")
        for tag in ("v", "is"):
            for no in c.findall(f"{{{NS}}}{tag}"):
                c.remove(no)
        c.attrib.pop("t", None)  # resultado numérico
        if not e["mesma"]:
            if f is None:
                f = etree.SubElement(c, f"{{{NS}}}f")
            for atributo in ("t", "si", "ref"):
                f.attrib.pop(atributo, None)
            f.text = e["formula"].lstrip("=")
        v = etree.SubElement(c, f"{{{NS}}}v")
        v.text = repr(e["valor"])
    return etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)


def _patch_workbook(xml):
    """Pede ao Excel para recalcular tudo ao abrir (os dependentes têm valor em cache antigo)."""
    raiz = etree.fromstring(xml)
    calc = raiz.find(f"{{{NS}}}calcPr")
    if calc is not None:
        calc.set("fullCalcOnLoad", "1")
    return etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)


def corrigir_planilha(caminho_origem, meta, correcoes, pasta_destino=None):
    """Aplica `correcoes` (de sanitizacao.sanitizar) na aba do mês e salva `<nome>_Corrigida.xlsx`.

    meta: o 4º retorno de ler_producao. Retorna (Path do arquivo, lista de edições) ou (None, [])
    se não havia correção que a planilha precise.
    """
    origem = Path(caminho_origem)
    correcoes = [c for c in correcoes if c["chave"] in meta["celulas"]]
    if not correcoes:
        return None, []
    aba = meta["aba"]
    edicoes = _planejar(origem, aba, meta["col_rotulo"], meta["celulas"], correcoes)

    destino = Path(pasta_destino or origem.parent) / f"{origem.stem}{SUFIXO}{origem.suffix}"
    with zipfile.ZipFile(origem) as zin:
        caminho_aba = _caminho_da_aba(zin, aba)
        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                dados = zin.read(item.filename)
                if item.filename == caminho_aba:
                    dados = _patch_aba(dados, edicoes)
                elif item.filename == "xl/workbook.xml":
                    dados = _patch_workbook(dados)
                elif item.filename == "xl/calcChain.xml":
                    continue  # lista de cálculo antiga; o Excel recria
                elif item.filename == "[Content_Types].xml":
                    dados = re.sub(rb'<Override[^>]*calcChain[^>]*/>', b"", dados)
                elif item.filename == "xl/_rels/workbook.xml.rels":
                    dados = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b"", dados)
                zout.writestr(item, dados)
    return destino, [{"celula": f"{aba}!{k}", **v} for k, v in edicoes.items()]
