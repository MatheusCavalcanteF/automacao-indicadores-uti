"""Write-back: corrige as fórmulas quebradas na planilha de Produção e salva `<nome>_Corrigida_<MES>.xlsx`.

Leitura e conferência com openpyxl; a escrita é um patch direto no XML da célula. Motivo: um
load/save do openpyxl reescreve o arquivo inteiro e descarta formas/caixas de texto e parte da
formatação dos gráficos (a planilha real tem 20 desenhos e 28 gráficos). O patch troca só a
fórmula (e o valor em cache) das células corrigidas; o resto do arquivo fica idêntico.
O original nunca é alterado. O mês vai no nome para que rodar outro mês não apague a cópia de auditoria.
"""

import os
import re
import zipfile
from copy import deepcopy
from pathlib import Path

import openpyxl
from lxml import etree
from openpyxl.utils import get_column_letter

from utils import SUFIXO_CORRIGIDA, resolver_mes

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


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


def _pulada(aba, chave, celulas, motivo):
    celula = f"{aba}!{_addr(celulas, chave)}" if chave in celulas else f"{aba}!?"
    return {"chave": chave, "celula": celula, "motivo": motivo}


def _planejar(caminho, aba, col_rotulo, celulas, correcoes):
    """Confere cada célula-alvo (rótulo na coluna A etc.) e monta as edições.

    Retorna (edicoes {addr: dict}, puladas [dict]). Correção que não passa na conferência é pulada,
    com o motivo, sem impedir as demais.
    """
    wb = openpyxl.load_workbook(caminho)  # só leitura: nunca salvamos este objeto
    ws = wb[aba]
    edicoes, puladas = {}, []
    for corr in correcoes:
        chave = corr["chave"]
        try:
            if chave not in celulas:
                raise ValueError(f"indicador '{chave}' sem célula rastreada.")
            info = celulas[chave]
            rotulo = ws.cell(info["linha"], col_rotulo).value
            if str(rotulo or "").strip() != info["rotulo"].strip():
                raise ValueError(f"linha {info['linha']}: rótulo '{rotulo}' ≠ esperado '{info['rotulo']}' "
                                 f"(planilha mudou desde a leitura?).")
            endereco = _addr(celulas, chave)
            atual = ws[endereco].value
            nova = _formula(corr["acao"], chave, celulas)
        except KeyError as e:
            puladas.append(_pulada(aba, chave, celulas, f"a fórmula depende de {e}, que não tem célula rastreada."))
            continue
        except (ValueError, TypeError) as e:
            puladas.append(_pulada(aba, chave, celulas, str(e)))
            continue
        mesma = isinstance(atual, str) and _canonica(atual) == _canonica(nova)
        edicoes[endereco] = {
            "chave": chave, "atual": atual, "formula": nova, "valor": float(corr["novo"]), "mesma": mesma,
        }
    return edicoes, puladas


def _caminho_da_aba(zf, aba):
    wb = etree.fromstring(zf.read("xl/workbook.xml"))
    rid = next((s.get(f"{{{NS_REL}}}id") for s in wb.iter(f"{{{NS}}}sheet") if s.get("name") == aba), None)
    if rid is None:
        raise ValueError(f"Aba '{aba}' não encontrada no workbook.xml.")
    rels = etree.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    alvo = next(r.get("Target") for r in rels if r.get("Id") == rid)
    return alvo.lstrip("/") if alvo.startswith("/") else "xl/" + alvo


def _aplicar_na_celula(c, endereco, e):
    """Troca fórmula e valor em cache de um <c>. Conferências antes de qualquer alteração."""
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


def _patch_aba(xml, edicoes):
    """Aplica as edições no XML da aba. Retorna (xml, {addr: motivo} das células puladas).

    Cada célula é independente: se uma falha, ela volta ao estado original e as demais seguem.
    """
    raiz = etree.fromstring(xml)
    celulas = {c.get("r"): c for c in raiz.iter(f"{{{NS}}}c")}
    puladas = {}
    for endereco, e in edicoes.items():
        c = celulas.get(endereco)
        if c is None:
            puladas[endereco] = f"célula {endereco} não existe no XML da aba."
            continue
        original = deepcopy(c)
        try:
            _aplicar_na_celula(c, endereco, e)
        except Exception as erro:  # noqa: BLE001 - uma célula com problema não derruba as demais
            c.getparent().replace(c, original)
            puladas[endereco] = str(erro)
    return etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True), puladas


def _patch_workbook(xml):
    """Pede ao Excel para recalcular tudo ao abrir (os dependentes têm valor em cache antigo)."""
    raiz = etree.fromstring(xml)
    calc = raiz.find(f"{{{NS}}}calcPr")
    if calc is not None:
        calc.set("fullCalcOnLoad", "1")
    return etree.tostring(raiz, xml_declaration=True, encoding="UTF-8", standalone=True)


def nome_da_copia(origem, mes):
    """'PLANILHA X.xlsx', 'set' -> 'PLANILHA X_Corrigida_SETEMBRO.xlsx'."""
    origem = Path(origem)
    return f"{origem.stem}{SUFIXO_CORRIGIDA}_{resolver_mes(mes)}{origem.suffix}"


def corrigir_planilha(caminho_origem, meta, correcoes, mes, pasta_destino=None):
    """Aplica `correcoes` (de sanitizacao.sanitizar) na aba do mês e salva `<nome>_Corrigida_<MES>.xlsx`.

    meta: o 4º retorno de ler_producao. Retorna (Path do arquivo ou None, edições aplicadas, puladas).
    puladas: [{"chave", "celula", "motivo"}] — células que não puderam ser corrigidas e devem ser
    ajustadas à mão; as demais correções são gravadas normalmente. Sem nenhuma edição aplicável,
    nada é gravado (Path = None).

    A cópia é escrita num arquivo temporário e só então renomeada: uma falha no meio não deixa um
    `_Corrigida_<MES>.xlsx` corrompido. O arquivo de origem nunca é aberto para escrita.
    """
    origem = Path(caminho_origem)
    correcoes = [c for c in correcoes if c["chave"] in meta["celulas"]]
    if not correcoes:
        return None, [], []
    aba = meta["aba"]
    edicoes, puladas = _planejar(origem, aba, meta["col_rotulo"], meta["celulas"], correcoes)
    if not edicoes:
        return None, [], puladas

    destino = Path(pasta_destino or origem.parent) / nome_da_copia(origem, mes)
    if destino.resolve() == origem.resolve():
        raise ValueError(f"Destino igual à origem ({origem.name}); a planilha original não pode ser sobrescrita.")
    temporario = destino.with_name(destino.name + ".tmp")  # extensão .tmp: localizar_arquivo ignora
    puladas_xml = {}
    try:
        with zipfile.ZipFile(origem) as zin:
            caminho_aba = _caminho_da_aba(zin, aba)
            with zipfile.ZipFile(temporario, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    dados = zin.read(item.filename)
                    if item.filename == caminho_aba:
                        dados, puladas_xml = _patch_aba(dados, edicoes)
                    elif item.filename == "xl/workbook.xml":
                        dados = _patch_workbook(dados)
                    elif item.filename == "xl/calcChain.xml":
                        continue  # lista de cálculo antiga; o Excel recria
                    elif item.filename == "[Content_Types].xml":
                        dados = re.sub(rb'<Override[^>]*calcChain[^>]*/>', b"", dados)
                    elif item.filename == "xl/_rels/workbook.xml.rels":
                        dados = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b"", dados)
                    zout.writestr(item, dados)

        for endereco, motivo in puladas_xml.items():
            puladas.append({"chave": edicoes.pop(endereco)["chave"], "celula": f"{aba}!{endereco}", "motivo": motivo})
        if not edicoes:
            return None, [], puladas
        try:
            os.replace(temporario, destino)
        except PermissionError as e:
            raise PermissionError(f"Feche o arquivo se estiver aberto no Excel: {destino}") from e
    finally:
        temporario.unlink(missing_ok=True)
    return destino, [{"celula": f"{aba}!{k}", **v} for k, v in edicoes.items()], puladas
