"""Pipeline mensal: extrai as fontes, sanitiza, corrige a planilha, valida e gera o relatório Word.

Uso: python main.py MES [ANO]      (ANO padrão: ano atual)
Exemplo: python main.py OUTUBRO 2026

Funciona com os arquivos que existirem em dados_entrada/ (só a planilha de Produção Diária é
indispensável). Censo, PDF da Fisioterapia ou ODT de Testes ausentes — ou existentes só de outro
mês — geram [WARNING] e as variáveis ficam "[Sem Dados]"; o relatório sai normalmente.
Produção é a fonte da verdade: divergências com o Censo são resolvidas a favor dela ([INFO]),
fórmulas quebradas são recalculadas ([CORREÇÃO]) e gravadas numa cópia da planilha
(<nome>_Corrigida.xlsx, [EXCEL]); o original nunca é alterado.
"""

import argparse
import sys
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ / "scripts"))

from consolidacao import consolidar_e_sanitizar, extrair_por_fonte  # noqa: E402
from corrigir_planilha import corrigir_planilha  # noqa: E402
from exportar_mapa import exportar_mapa  # noqa: E402
from gerador_relatorio import criar_template_exemplo, gerar_relatorio_word  # noqa: E402
from localizar_arquivos import resolver_mes  # noqa: E402
from validacao import validar  # noqa: E402

TEMPLATE = RAIZ / "templates" / "modelo_relatorio.docx"
MAPA = RAIZ / "mapa_variaveis.txt"


def _corrigir_excel(auxiliar, correcoes):
    """Grava as correções na cópia da planilha. Falha aqui nunca impede o relatório."""
    if not correcoes:
        return
    try:
        destino, edicoes = corrigir_planilha(auxiliar["arquivo_producao"], auxiliar["meta"], correcoes)
    except Exception as e:  # noqa: BLE001 - qualquer problema vira aviso; o Word não depende disso
        print(f"[EXCEL][WARNING] Planilha NÃO corrigida: {e}")
        return
    for ed in edicoes:
        if ed["mesma"]:
            print(f"[EXCEL] {ed['celula']}: fórmula {ed['atual']} já correta; valor atualizado para {ed['valor']:.6g}.")
        else:
            print(f"[EXCEL] {ed['celula']}: {ed['atual']} -> {ed['formula']} (= {ed['valor']:.6g})")
    print(f"[EXCEL] Planilha de Produção recalculada e salva com correções: {destino.name}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mes", help="mês de referência (ex.: SETEMBRO, set, 9)")
    ap.add_argument("ano", nargs="?", default=str(date.today().year))
    args = ap.parse_args()
    try:
        mes, ano = resolver_mes(args.mes), args.ano
    except ValueError as e:
        sys.exit(str(e))

    print(f"Referência: {mes}/{ano}")
    try:
        secoes, auxiliar = extrair_por_fonte(mes, ano)
    except (FileNotFoundError, ValueError) as e:
        sys.exit(f"Erro na extração (a planilha de Produção é indispensável): {e}")

    dados, log, correcoes = consolidar_e_sanitizar(secoes, auxiliar["series"])
    print(f"Total de variáveis consolidadas: {len(dados)}")
    for nivel, msg in log:
        print(f"[{nivel}] {msg}")

    _corrigir_excel(auxiliar, correcoes)

    extras = {"mes": mes, "ano": ano}
    secoes_finais = {titulo: {k: dados[k] for k in valores} for titulo, valores in secoes.items()}
    exportar_mapa(secoes_finais, MAPA, extras=extras)  # gabarito = exatamente o que vai ao Word

    for nivel, msg in validar(dados, auxiliar["series"], auxiliar["titulo_aba"], mes, ano):
        print(f"[{nivel}] {msg}")

    if not TEMPLATE.is_file():
        print(f"[WARNING] {TEMPLATE.name} não existe; criando modelo de exemplo (rode scripts/montar_template.py).")
        criar_template_exemplo(TEMPLATE, dados)

    saida = RAIZ / "relatorios_gerados" / f"Relatorio_{mes.capitalize()}_{ano}_Final.docx"
    arquivo, faltantes = gerar_relatorio_word(dados, TEMPLATE, saida, extras=extras, series=auxiliar["series"])
    print("Relatório salvo em:", arquivo)
    if faltantes:
        print(f"[WARNING] {len(faltantes)} variável(is) do modelo sem valor ('[Sem Dados]' no Word): {faltantes}")
    if auxiliar["pendencias"]:
        print(f"[WARNING] Fontes sem dados: {', '.join(auxiliar['pendencias'])}.")
    print("Campos manuais: busque '[PREENCHER' (amarelo) no Word.")


if __name__ == "__main__":
    main()
