"""Preenche o modelo Word do Relatório de Avaliação dos Indicadores com docxtpl.

O contexto leva números crus; a formatação é feita no template por filtros Jinja:
  {{ taxa_de_ocupacao_global | pct }}   -> 93,67%
  {{ n_de_saidas | num }}               -> 29 (inteiro sem casas; senão 2 casas, vírgula)
  {{ media_de_permanencia | dec }}      -> 7,80
  {{ serie__x | dia(3) }}               -> valor do 3º dia da série 'x'
Variável inexistente não quebra a geração: aparece como [Sem Dados] e é listada.
Valores já pendentes (strings) passam pelos filtros sem alteração.
"""

from pathlib import Path

import jinja2
from docx import Document
from docxtpl import DocxTemplate

from extracao_producao import chave_variavel, extrair_dados_diarios

RAIZ = Path(__file__).resolve().parent.parent

chave_template = chave_variavel


SEM_DADOS = "[Sem Dados]"


class _Faltante(jinja2.Undefined):
    """Undefined que registra o nome da variável e deixa um marcador visível no Word."""

    registro = set()

    def __str__(self):
        _Faltante.registro.add(self._undefined_name)
        return SEM_DADOS

    __html__ = __str__


def _br(valor, casas):
    return f"{valor:.{casas}f}".replace(".", ",")


def _numerico(valor):
    return isinstance(valor, (int, float)) and not isinstance(valor, bool)


def _texto(valor):
    """Não numérico (marcador de pendência, variável ausente, None) passa direto, sem quebrar."""
    return "-" if valor is None else str(valor)


def pct(valor, casas=2):
    """0.9367 -> '93,67%'."""
    return _br(valor * 100, casas) + "%" if _numerico(valor) else _texto(valor)


def dec(valor, casas=2):
    return _br(valor, casas) if _numerico(valor) else _texto(valor)


def num(valor):
    """Inteiro sem casas; fracionário com 2 casas."""
    if not _numerico(valor):
        return _texto(valor)
    return str(int(valor)) if float(valor).is_integer() else _br(valor, 2)


def dia(serie, indice):
    """Valor do dia `indice` (1 = primeiro dia) de uma série; vazio se o mês não tem esse dia."""
    if not isinstance(serie, (list, tuple)) or not 1 <= indice <= len(serie) or serie[indice - 1] is None:
        return ""
    return num(serie[indice - 1])


def _ambiente():
    env = jinja2.Environment(undefined=_Faltante, autoescape=False)
    env.filters.update(pct=pct, dec=dec, num=num, fmt=num, dia=dia)
    return env


def montar_contexto(dados, extras=None, series=None):
    """{variável: número} + extras (mês, ano...) + séries diárias como 'serie__<variável>'."""
    contexto = dict(dados)
    for chave, valores in (series or {}).items():
        contexto[f"serie__{chave}"] = valores
    contexto.update(extras or {})
    return contexto


def criar_template_exemplo(caminho_template, dados):
    """Modelo mínimo com todas as variáveis, só para testes quando não há modelo real."""
    doc = Document()
    doc.add_heading("Relatório de Avaliação dos Indicadores - UTI Pediátrica (MODELO DE EXEMPLO)", level=1)
    doc.add_paragraph("Referência: {{ mes }}/{{ ano }}")
    for nome in dados:
        filtro = "pct" if nome.startswith("taxa") and "__" not in nome else "num"
        doc.add_paragraph(f"{nome}: {{{{ {nome} | {filtro} }}}}")
    Path(caminho_template).parent.mkdir(parents=True, exist_ok=True)
    doc.save(caminho_template)


def gerar_relatorio_word(dados, caminho_template, caminho_saida, extras=None, series=None):
    """Renderiza o template e salva em caminho_saida. Retorna (Path salvo, variáveis faltantes)."""
    caminho_template = Path(caminho_template)
    if not caminho_template.is_file():
        raise FileNotFoundError(f"Template não encontrado: {caminho_template}")

    _Faltante.registro.clear()
    try:
        modelo = DocxTemplate(caminho_template)
        modelo.render(montar_contexto(dados, extras, series), _ambiente())
    except Exception as e:
        raise ValueError(f"Falha ao renderizar '{caminho_template.name}': {e}") from e

    caminho_saida = Path(caminho_saida)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)
    try:
        modelo.save(caminho_saida)
    except PermissionError as e:
        raise PermissionError(f"Feche o arquivo se estiver aberto no Word: {caminho_saida}") from e
    return caminho_saida, sorted(_Faltante.registro)


if __name__ == "__main__":
    import sys

    from consolidacao import consolidar_e_sanitizar, extrair_por_fonte

    if len(sys.argv) < 3:
        sys.exit("Uso: python scripts/gerador_relatorio.py MES ANO")
    mes, ano = sys.argv[1].upper(), sys.argv[2]
    secoes, auxiliar = extrair_por_fonte(mes, ano)
    dados, _, _ = consolidar_e_sanitizar(secoes, auxiliar["series"])
    template = RAIZ / "templates" / "modelo_relatorio.docx"
    saida = RAIZ / "relatorios_gerados" / f"Relatorio_{mes.capitalize()}_{ano}_Teste.docx"
    arquivo, faltantes = gerar_relatorio_word(
        dados, template, saida, extras={"mes": mes, "ano": ano}, series=auxiliar["series"]
    )
    print("Relatório salvo em:", arquivo, "| variáveis ausentes:", faltantes)
