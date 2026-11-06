"""Gráficos anuais: valores das abas mensais, desenho e troca da imagem no Word. Tudo sintético, em tmp_path."""

import hashlib
import io
import zipfile

import openpyxl
import pytest
from docx import Document
from docx.oxml import parse_xml
from PIL import Image

import graficos as g
import main
from extracao_producao import ler_producao
from gerador_relatorio import gerar_relatorio_word
from gerar_dados_exemplo import gerar
from montar_template import TABELAS, montar
from sanitizacao import SEM_DADOS, sanitizar
from utils import resolver_periodo

OCUPACAO = "taxa_de_ocupacao_global"
HM = "taxa_de_adesao_ao_protocolo_de_hm_momentos"
SAIDAS = "saidas_hospitalares__n_de_saidas"
NS_VML = 'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
NS_WR = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


# ---------------------------------------------------------------- planilha sintética

def _aba(wb, nome, pacientes, leitos=10, hm=None):
    """Aba no layout da Produção: taxa = pacientes / leitos, por dia; 3 dias bastam para o extrator."""
    ws = wb.create_sheet(nome)
    ws.append(["UTI PEDIÁTRICA - DADOS SINTÉTICOS"])
    ws.append(["DATAS", 1, 2, 3, "TOTAL", "REALIZADO"])
    total = pacientes * 3
    ws.append(["PACIENTE DIA GLOBAL", pacientes, pacientes, pacientes, total, None])
    ws.append(["LEITO DIA GLOBAL", leitos, leitos, leitos, leitos * 3, None])
    ws.append([])
    ws.append(["SAÍDAS HOSPITALARES", None, None, None, "TOTAL", 10])
    ws.append(["N. DE SAÍDAS", 3, 3, 4, 10, None])
    ws.append([])
    ws.append(["TAXA DE OCUPAÇÃO GLOBAL", None, None, None, "TOTAL", pacientes / leitos])
    ws.append(["N. DE PACIENTES DIA", pacientes, pacientes, pacientes, total, None])
    ws.append(["N. DE LEITO DIAS", leitos, leitos, leitos, leitos * 3, None])
    ws.append([])
    dias = [hm, hm, hm] if hm is not None else [None, None, None]
    ws.append(["% TAXA DE ADESÃO AO PROTOCOLO DE HM - MOMENTOS", None, None, None, "TOTAL", 0])
    ws.append(["N. DE OBSERVAÇÕES", *dias, sum(d or 0 for d in dias), None])


@pytest.fixture
def planilha(tmp_path):
    """Janeiro, março, abril e dezembro; fevereiro sem aba. O FECHAMENTO é um chamariz: valores absurdos."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _aba(wb, "UTI PEDIÁTRICA JANEIRO", 5, hm=9)
    _aba(wb, "UTI PEDIÁTRICA MARÇO", 6)
    _aba(wb, "UTI PEDIÁTRICA ABRIL", 7)
    _aba(wb, "UTI PEDIÁTRICA DEZ", 8)
    fechamento = wb.create_sheet("FECHAMENTO COMPILADO ANUAL")
    fechamento.append(["TAXA DE OCUPAÇÃO GLOBAL", "JANEIRO", "MARÇO"])
    fechamento.append([2025, "#REF!", 99])
    caminho = tmp_path / "PLANILHA PRODUCAO 2026.xlsx"
    wb.save(caminho)
    return caminho


def _dados_do_mes(planilha, mes):
    totais, series, _, _ = ler_producao(planilha, mes)
    return sanitizar(totais, series)[0]


def test_historico_vem_das_abas_mensais_e_para_no_mes_de_referencia(planilha):
    periodo = resolver_periodo("abril", 2026)
    historico = g.historico_anual(planilha, periodo, _dados_do_mes(planilha, "abril"))

    assert historico.ano == 2026 and historico.n_meses == 4
    # Fevereiro não tem aba; dezembro existe, mas é depois da referência; o FECHAMENTO (99) não entra.
    assert historico.series[OCUPACAO] == pytest.approx([0.5, None, 0.6, 0.7] + [None] * 8, nan_ok=False)
    assert set(historico.series) == set(g.GRAFICOS)


def test_mes_de_referencia_usa_os_dados_do_relatorio(planilha):
    dados = _dados_do_mes(planilha, "abril") | {OCUPACAO: 0.77}  # ex.: valor recalculado pela sanitização
    historico = g.historico_anual(planilha, resolver_periodo("abril", 2026), dados)

    assert historico.series[OCUPACAO][3] == 0.77


def test_janeiro_nao_precisa_de_abas_anteriores(planilha):
    historico = g.historico_anual(planilha, resolver_periodo("janeiro", 2026), {OCUPACAO: 0.5})

    assert historico.series[OCUPACAO] == [0.5] + [None] * 11


def test_sem_dados_e_valor_invalido_nao_viram_zero(planilha):
    dados = {OCUPACAO: "[Valor Inválido na Planilha]", HM: SEM_DADOS}
    historico = g.historico_anual(planilha, resolver_periodo("abril", 2026), dados)

    assert historico.series[OCUPACAO][3] is None
    assert historico.series[HM][2:4] == [None, None]  # março: bloco sem nenhum dia preenchido


# ---------------------------------------------------------------- metas e rótulos

@pytest.mark.parametrize("texto, filtro, esperado", [
    ("40", "num", 40.0),
    (">=80%", "pct", 0.8),
    ("<=6,5 dias", "dec", 6.5),
    ("< 0,25", "pct", 0.0025),
    ("41,27%", "pct", 0.4127),
    ("3,50", "dec", 3.5),
    ("0", "dec", 0.0),
    ("sem número", "pct", None),
])
def test_meta_lida_do_contratado(texto, filtro, esperado):
    assert g.meta_do_contratado(texto, filtro) == esperado


def test_prazo_no_contratado_nao_vira_meta():
    chave = "atender_100_das_solicitacoes_de_interconsultas_em_neuro"
    assert g.meta_do_contratado("Em até 12h", "pct", chave) == 1.0


def test_todo_indicador_do_modelo_tem_grafico():
    assert {r[1] for r in TABELAS} == set(g.GRAFICOS)


@pytest.mark.parametrize("valor, formato, casas, esperado", [
    (0, "pct", 2, "0%"),
    (0.0712, "pct", 2, "7,12%"),
    (0.7349, "pct", 0, "73%"),
    (1.0, "pct", 0, "100%"),
    (5.26, "dec", 1, "5,3"),
    (31.0, "dec", 0, "31"),
])
def test_rotulo_da_barra(valor, formato, casas, esperado):
    assert g.rotulo_valor(valor, g.Grafico("T", formato, casas)) == esperado


@pytest.mark.parametrize("valor, formato, esperado", [
    (0.8, "pct", "80%"), (0.4127, "pct", "41,27%"), (6.5, "dec", "6,5"), (40.0, "dec", "40"), (0.0, "dec", "0"),
])
def test_rotulo_da_meta(valor, formato, esperado):
    assert g.rotulo_meta(valor, g.Grafico("T", formato, 0)) == esperado


# ---------------------------------------------------------------- desenho

def _textos(figura):
    eixo = figura.axes[0]
    return [t.get_text() for t in eixo.texts]


def test_figura_tem_doze_meses_barras_e_meta():
    valores = [0.5, None, 0.6] + [None] * 9
    figura = g.figura_grafico(g.GRAFICOS[OCUPACAO], valores, 3, 2026, 0.9)
    eixo = figura.axes[0]

    assert [t.get_text() for t in eixo.get_xticklabels()] == list(g.MESES_ROTULO)
    assert [p.get_height() for p in eixo.patches] == [0.5, 0.6]
    assert [p.get_x() + p.get_width() / 2 for p in eixo.patches] == pytest.approx([0, 2])
    assert _textos(figura).count(g.SEM_DADO) == 1  # só fevereiro: os meses futuros ficam vazios
    assert {"50%", "60%", "90%"} <= set(_textos(figura))
    assert list(eixo.lines[0].get_ydata()) == [0.9, 0.9]
    assert eixo.get_ylim()[0] == 0
    assert [t.get_text() for t in figura.legends[0].get_texts()] == ["2026", "META"]
    assert figura.axes[0].get_title() == g.GRAFICOS[OCUPACAO].titulo


def test_mes_sem_dado_nao_desenha_barra_de_zero():
    figura = g.figura_grafico(g.GRAFICOS[HM], [None] * 12, 9, 2026, 0.95)

    assert not figura.axes[0].patches
    assert _textos(figura).count(g.SEM_DADO) == 9
    assert "0%" not in _textos(figura)
    legenda = figura.legends[0].get_patches()[0]
    assert legenda.get_facecolor()[:3] == pytest.approx((1.0, 0xC0 / 255, 0.0))  # não a cor padrão


def test_nota_do_sem_dado_aparece_na_imagem():
    """A nota é desenhada no canto inferior direito, dentro da imagem."""
    grafico = g.GRAFICOS[HM]
    com = Image.open(io.BytesIO(g.png_do_grafico(grafico, [None] * 12, 1, 2026, 0.95))).convert("L")
    sem = Image.open(io.BytesIO(g.png_do_grafico(grafico, [0.5] + [None] * 11, 1, 2026, 0.95))).convert("L")
    largura, altura = com.size
    canto = (largura - 300, altura - 50, largura, altura)

    assert min(com.crop(canto).tobytes()) < 200
    assert min(sem.crop(canto).tobytes()) == 255


def test_todos_os_zeros_e_sem_meta_nao_quebram():
    figura = g.figura_grafico(g.GRAFICOS["evento_sentinela_0"], [0.0] * 3 + [None] * 9, 3, 2026)
    assert figura.axes[0].get_ylim() == (0, 1)
    assert not figura.axes[0].lines


def test_png_tem_o_tamanho_da_imagem_original():
    png = g.png_do_grafico(g.GRAFICOS[OCUPACAO], [0.5] * 12, 12, 2026, 0.9, tamanho=(752, 452))
    assert Image.open(io.BytesIO(png)).size == (752 * g.DPI // 100, 452 * g.DPI // 100)


# ---------------------------------------------------------------- Word

def _png(cor):
    saida = io.BytesIO()
    Image.new("RGB", (752, 434), cor).save(saida, format="PNG")
    return saida


def _paragrafo_vml(doc, cor):
    rid, _ = doc.part.get_or_add_image(_png(cor))
    p = doc.add_paragraph()
    p._p.append(parse_xml(
        f'<w:r {NS_WR}><w:pict {NS_VML}><v:shape id="s{rid}" type="#_x0000_t75" style="width:360pt;height:208pt">'
        f'<v:imagedata r:id="{rid}" o:title=""/></v:shape></w:pict></w:r>'))


def _tabela_indicador(doc, titulo, contratado, realizado):
    tabela = doc.add_table(rows=3, cols=3)
    tabela.rows[0].cells[1].text = titulo
    for celula, texto in zip(tabela.rows[1].cells, ("", "Contratado", "Realizado")):
        celula.text = texto
    for celula, texto in zip(tabela.rows[2].cells, ("UTI PEDIÁTRICA", contratado, realizado)):
        celula.text = texto


@pytest.fixture
def base(tmp_path):
    """Capa, duas tabelas de indicador com imagem logo depois, e os valores 'do mês antigo' nas imagens."""
    doc = Document()
    _paragrafo_vml(doc, "gray")  # capa, antes de qualquer indicador
    _tabela_indicador(doc, "% Taxa de Ocupação ≥", ">=80%", "70%")
    _paragrafo_vml(doc, "red")
    _tabela_indicador(doc, "Saídas", "40", "20")
    _paragrafo_vml(doc, "blue")
    caminho = tmp_path / "base.docx"
    doc.save(caminho)
    return caminho


def _alts(caminho):
    corpo = Document(caminho).element.body
    return [f.get("alt") for f in corpo.iter(f"{g._NS_VML}shape")]


def _hashes_de_imagens(caminho):
    with zipfile.ZipFile(caminho) as z:
        return {hashlib.sha1(z.read(n)).hexdigest() for n in z.namelist() if n.startswith("word/media/")}


def _sha1(buffer):
    return hashlib.sha1(buffer.getvalue()).hexdigest()


def test_montar_marca_cada_imagem_com_indicador_e_meta_e_poupa_a_capa(base, tmp_path):
    modelo = montar(base, tmp_path / "modelo.docx")

    assert _alts(modelo) == [None, f"grafico:{OCUPACAO};meta=0.8", f"grafico:{SAIDAS};meta=40"]


def test_imagem_sem_tabela_antes_e_tabela_sem_imagem_sao_avisadas(tmp_path, capsys):
    doc = Document()
    _tabela_indicador(doc, "Saídas", "40", "20")
    _paragrafo_vml(doc, "red")
    _paragrafo_vml(doc, "blue")  # segunda imagem: sem indicador
    _tabela_indicador(doc, "% Taxa de Ocupação ≥", ">=80%", "70%")  # sem imagem
    caminho = tmp_path / "base.docx"
    doc.save(caminho)

    montar(caminho, tmp_path / "modelo.docx")

    saida = capsys.readouterr().out
    assert "1 imagem(ns) sem tabela" in saida
    assert f"sem imagem de gráfico no modelo-base: ['{OCUPACAO}']" in saida


def _historico(ocupacao=0.5, saidas=14.0):
    series = {chave: [None] * 12 for chave in g.GRAFICOS}
    series[OCUPACAO] = [ocupacao] + [None] * 11
    series[SAIDAS] = [saidas] + [None] * 11
    return g.HistoricoAnual(2026, 1, series)


def test_gerador_troca_as_imagens_e_descarta_as_antigas(base, tmp_path):
    modelo = montar(base, tmp_path / "modelo.docx")
    saida, _, _ = gerar_relatorio_word({}, modelo, tmp_path / "r.docx", graficos=_historico())

    imagens = _hashes_de_imagens(saida)
    assert _sha1(_png("red")) not in imagens and _sha1(_png("blue")) not in imagens
    assert _sha1(_png("gray")) in imagens  # capa
    assert len(imagens) == 3
    alts = _alts(saida)
    assert alts[0] is None
    assert not any("grafico:" in (a or "") for a in alts)
    assert alts[1] == f"Gráfico: {g.GRAFICOS[OCUPACAO].titulo}"


def test_relatorios_de_meses_diferentes_tem_imagens_diferentes(base, tmp_path):
    modelo = montar(base, tmp_path / "modelo.docx")
    um, _, _ = gerar_relatorio_word({}, modelo, tmp_path / "a.docx", graficos=_historico(0.5))
    outro, _, _ = gerar_relatorio_word({}, modelo, tmp_path / "b.docx", graficos=_historico(0.8))

    assert _hashes_de_imagens(um) != _hashes_de_imagens(outro)


def test_meses_sem_dado_entram_em_conferir_e_sem_historico_nao_fica_grafico_antigo(base, tmp_path):
    modelo = montar(base, tmp_path / "modelo.docx")
    historico = _historico()._replace(n_meses=2)
    saida, _, _ = gerar_relatorio_word({}, modelo, tmp_path / "r.docx", graficos=historico,
                                       avisos={"PENDENCIA": ["forca a página"]})

    textos = [p.text for p in Document(saida).paragraphs]
    assert any(g.GRAFICOS[OCUPACAO].titulo in t and "Fevereiro" in t for t in textos)

    saida, _, _ = gerar_relatorio_word({}, modelo, tmp_path / "s.docx")  # sem graficos
    imagens = _hashes_de_imagens(saida)
    assert _sha1(_png("red")) not in imagens and _sha1(_png("blue")) not in imagens


def test_imagem_inline_tambem_e_trocada(tmp_path):
    """Documento-base salvo pelo Word moderno: wp:inline/a:blip em vez de VML."""
    doc = Document()
    _tabela_indicador(doc, "Saídas", "40", "20")
    doc.add_paragraph().add_run().add_picture(_png("red"))
    caminho = tmp_path / "base.docx"
    doc.save(caminho)

    modelo = montar(caminho, tmp_path / "modelo.docx")
    saida, _, _ = gerar_relatorio_word({}, modelo, tmp_path / "r.docx", graficos=_historico())

    assert _sha1(_png("red")) not in _hashes_de_imagens(saida)


def test_modelo_antigo_sem_marcadores_e_detectado(base, tmp_path):
    assert "montar_template" in g.aviso_modelo_sem_marcadores(base)
    assert g.aviso_modelo_sem_marcadores(montar(base, tmp_path / "modelo.docx")) is None
    doc = Document()
    doc.add_paragraph("sem imagens")
    doc.save(tmp_path / "exemplo.docx")
    assert g.aviso_modelo_sem_marcadores(tmp_path / "exemplo.docx") is None


# ---------------------------------------------------------------- ponta a ponta (dados fictícios)

def _rodar(entrada, saida, *extra):
    main.main(["setembro", "2026", "--pasta-entrada", str(entrada), "--pasta-saida", str(saida), *extra])


def test_ponta_a_ponta_com_modelo_marcado(base, tmp_path, monkeypatch):
    monkeypatch.setattr(main, "TEMPLATE", montar(base, tmp_path / "modelo.docx"))
    entrada, saida = tmp_path / "entrada", tmp_path / "saida"
    gerar(entrada, "setembro", 2026)

    _rodar(entrada, saida)

    relatorio = saida / "Relatorio_Setembro_2026_Final.docx"
    imagens = _hashes_de_imagens(relatorio)
    assert _sha1(_png("red")) not in imagens and _sha1(_png("blue")) not in imagens
    textos = [p.text for p in Document(relatorio).paragraphs]
    # A planilha fictícia só tem a aba de setembro: os outros meses são avisados, não inventados.
    assert any(g.GRAFICOS[SAIDAS].titulo in t and "Janeiro" in t and "Setembro" not in t for t in textos)


def test_ponta_a_ponta_com_modelo_antigo_avisa_que_os_graficos_sao_velhos(base, tmp_path, monkeypatch):
    monkeypatch.setattr(main, "TEMPLATE", base)  # imagens sem marcador
    entrada, saida = tmp_path / "entrada", tmp_path / "saida"
    gerar(entrada, "setembro", 2026)

    _rodar(entrada, saida)

    textos = [p.text for p in Document(saida / "Relatorio_Setembro_2026_Final.docx").paragraphs]
    assert any("sem marcadores de gráfico" in t for t in textos)


def test_falha_ao_ler_as_abas_vira_erro_critico(base, tmp_path, monkeypatch):
    monkeypatch.setattr(main, "TEMPLATE", montar(base, tmp_path / "modelo.docx"))

    def falhar(*_):
        raise ValueError("aba corrompida")

    monkeypatch.setattr(main, "historico_anual", falhar)
    entrada, saida = tmp_path / "entrada", tmp_path / "saida"
    gerar(entrada, "setembro", 2026)

    _rodar(entrada, saida)

    relatorio = saida / "Relatorio_Setembro_2026_Final.docx"
    textos = [p.text for p in Document(relatorio).paragraphs]
    assert any("Gráficos não gerados" in t and "aba corrompida" in t for t in textos)
    assert _sha1(_png("red")) not in _hashes_de_imagens(relatorio)
