"""Parágrafos cujo texto vem do documento-base: só o número vira variável. Documentos sintéticos."""

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

from gerador_relatorio import gerar_relatorio_word
from montar_template import montar


def _base(tmp_path, *linhas):
    doc = Document()
    for linha in linhas:
        doc.add_paragraph(linha)
    caminho = tmp_path / "base.docx"
    doc.save(caminho)
    return caminho


def _textos(caminho):
    return [p.text for p in Document(caminho).paragraphs]


def test_nascidos_vivos_mantem_o_texto_do_documento_base(tmp_path):
    base = _base(tmp_path, "Tivemos um total de 57 nascidos vivos na Unidade Exemplo.")
    modelo = montar(base, tmp_path / "modelo.docx")

    assert _textos(modelo) == ["Tivemos um total de {{ n_de_nascidos_vivos | num }} nascidos vivos na Unidade Exemplo."]

    saida, faltantes, _ = gerar_relatorio_word({"n_de_nascidos_vivos": 61.0}, modelo, tmp_path / "r.docx")
    assert faltantes == []
    assert _textos(saida)[-1] == "Tivemos um total de 61 nascidos vivos na Unidade Exemplo."


def test_numero_com_separador_de_milhar_vira_uma_variavel_so(tmp_path):
    base = _base(tmp_path, "Tivemos um total de 1.234 nascidos vivos (Unidade Exemplo), 2 a mais que o mês anterior.")
    modelo = montar(base, tmp_path / "modelo.docx")

    assert _textos(modelo) == [
        "Tivemos um total de {{ n_de_nascidos_vivos | num }} nascidos vivos (Unidade Exemplo), 2 a mais que o mês anterior."
    ]


def _tabela_diaria(doc, linhas):
    tabela = doc.add_table(rows=len(linhas), cols=34)
    for r, valores in zip(tabela.rows, linhas):
        for cel, valor in zip(r.cells, valores):
            cel.text = valor
    return tabela


def test_tabelas_diarias_com_coluna_sobrando_e_sem_titulo(tmp_path):
    dias = [str(d) for d in range(1, 31)]
    doc = Document()
    _tabela_diaria(doc, [
        ["TITULO VM"] + [""] * 30 + ["TOTAL", "52%", ""],
        ["N. DE PACIENTES DIA EM VM"] + dias + ["134", "52%", ""],
        ["N. TOTAL DE PACIENTE DIA"] + dias + ["257", "52%", ""],
    ])
    _tabela_diaria(doc, [  # crônicos: sem linha de título e já com a coluna do dia 31
        ["PACIENTE DIA CRÔNICOS"] + dias + ["", "121", "41%"],
        ["N. TOTAL DE PACIENTE DIA"] + dias + ["", "260", "41%"],
    ])
    base = tmp_path / "base.docx"
    doc.save(base)

    modelo = Document(montar(base, tmp_path / "modelo.docx"))
    vm, cronicos = modelo.tables

    assert len(vm.columns) == len(cronicos.columns) == 34  # rótulo + 31 dias + TOTAL + taxa
    linha_vm = [c.text for c in vm.rows[1].cells]
    assert "dia(31)" in linha_vm[31]
    assert linha_vm[32] == "{{ taxa_de_utilizacao_de_ventilacao_mecanica__n_de_pacientes_dia_em_vm | num }}"
    assert linha_vm[33] == "{{ taxa_de_utilizacao_de_ventilacao_mecanica | pct }}"
    linha_cronicos = [c.text for c in cronicos.rows[0].cells]
    assert "serie__paciente_dia_cronicos" in linha_cronicos[1]
    assert linha_cronicos[32] == "{{ paciente_dia_cronicos | num }}"
    assert linha_cronicos[33] == "[PREENCHER]"


def test_frase_de_vm_que_abre_com_hifen_vira_variavel(tmp_path):
    base = _base(tmp_path, "-  Taxa de Utilização de VM total: tivemos uma taxa de 52,14%.")
    texto = _textos(montar(base, tmp_path / "modelo.docx"))[0]

    assert "{{ taxa_de_utilizacao_de_ventilacao_mecanica | pct(1) }}" in texto and "52,14" not in texto


def _linha_sumario(doc, texto, alvo, tipos=("begin", "end")):
    p = doc.add_paragraph(texto)
    run = p.add_run()._r
    run.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="{tipos[0]}"/>'))
    run.append(parse_xml(f'<w:instrText {nsdecls("w")}> PAGEREF {alvo} \\h </w:instrText>'))
    for tipo in tipos[1:]:
        run.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="{tipo}"/>'))
    return p


def test_sumario_perde_so_a_linha_sem_destino(tmp_path):
    doc = Document()
    _linha_sumario(doc, "abre o campo", "_TocSumido", tipos=("begin",))
    _linha_sumario(doc, "sem destino", "_TocSumido")
    _linha_sumario(doc, "com destino", "_TocOk")
    titulo = doc.add_paragraph("Título")
    titulo._p.insert(0, parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="1" w:name="_TocOk"/>'))
    base = tmp_path / "base.docx"
    doc.save(base)

    textos = _textos(montar(base, tmp_path / "modelo.docx"))

    assert "sem destino" not in textos
    assert "com destino" in textos and "abre o campo" in textos


def test_mes_do_documento_base_vira_variavel(tmp_path):
    base = _base(tmp_path, "SETEMBRO/2026", "referente ao mês de SETEMBRO/2026, demonstrando", "Convênio Nº. 03527/2024")
    modelo = montar(base, tmp_path / "modelo.docx")

    assert _textos(modelo) == [
        "{{ mes }}/{{ ano }}", "referente ao mês de {{ mes }}/{{ ano }}, demonstrando", "Convênio Nº. 03527/2024",
    ]

    saida, _, _ = gerar_relatorio_word({}, modelo, tmp_path / "r.docx", extras={"mes": "OUTUBRO", "ano": 2026})
    assert _textos(saida)[0] == "OUTUBRO/2026"


def test_paragrafo_sem_numero_fica_intacto_e_e_avisado(tmp_path, capsys):
    texto = "Tivemos um total de vários nascidos vivos na Unidade Exemplo."
    modelo = montar(_base(tmp_path, texto), tmp_path / "modelo.docx")

    assert _textos(modelo) == [texto]
    avisos = capsys.readouterr().out
    assert "parágrafos de base não encontrados" in avisos and "Tivemos um total de" in avisos
