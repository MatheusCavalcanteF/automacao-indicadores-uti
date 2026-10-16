"""Pipeline mensal: extrai as fontes, sanitiza, corrige a planilha, valida e gera o relatório Word.

Uso: python main.py MES [ANO] [--estrito] [--pasta-entrada PASTA] [--pasta-saida PASTA]
Exemplos: python main.py outubro 2026 | python main.py set 26 | python main.py 09/2026 --estrito
Dados fictícios: python scripts/gerar_dados_exemplo.py && python main.py setembro 2026
    --pasta-entrada exemplo/dados_entrada --pasta-saida exemplo/relatorios_gerados

Funciona com os arquivos que existirem na pasta de entrada (só a planilha de Produção Diária é
indispensável). Censo, PDF da Fisioterapia ou ODT de Testes ausentes — ou existentes só de outro
mês — geram [WARNING] e as variáveis ficam "[Sem Dados]".
Produção é a fonte da verdade: divergências com o Censo são resolvidas a favor dela ([INFO]),
fórmulas quebradas são recalculadas ([CORREÇÃO]) e gravadas numa cópia da planilha
(<nome>_Corrigida_<MES>.xlsx, [EXCEL]); o original nunca é alterado.

Erros não corrigíveis, fontes ausentes e pendências viram uma página inicial "AVISOS CRÍTICOS E
PENDÊNCIAS" no Word. Com --estrito, erro crítico ou fonte ausente impede a geração do Word.

Toda execução grava um log de auditoria na pasta de saída: execucao_<MES>_<ANO>_<AAAAMMDD_HHMMSS>.log,
com tudo o que sai no terminal e mais: quem rodou, em qual máquina e o hash de cada arquivo de entrada.

Códigos de saída: 0 = relatório gerado; 1 = falha (ex.: sem planilha de Produção);
2 = argumentos inválidos; 3 = bloqueado pelo modo estrito.
"""

import argparse
import getpass
import logging
import platform
import sys
from datetime import date
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ / "scripts"))

import log_execucao  # noqa: E402
from consolidacao import consolidar_e_sanitizar, extrair_por_fonte  # noqa: E402
from corrigir_planilha import corrigir_planilha  # noqa: E402
from exportar_mapa import exportar_mapa  # noqa: E402
from gerador_relatorio import (  # noqa: E402
    SECOES_AVISOS,
    criar_template_exemplo,
    dias_no_modelo,
    gerar_relatorio_word,
)
from utils import resolver_periodo  # noqa: E402
from validacao import validar  # noqa: E402

TEMPLATE = RAIZ / "templates" / "modelo_relatorio.docx"
MAPA = RAIZ / "mapa_variaveis.txt"
PASTA_ENTRADA = RAIZ / "dados_entrada"
PASTA_SAIDA = RAIZ / "relatorios_gerados"
SAIDA_FALHA, SAIDA_BLOQUEADA = 1, 3
BLOQUEIAM_NO_ESTRITO = ("CRITICO", "FONTE")
# Nível das mensagens de sanitização/validação no log (o rótulo entre colchetes continua no texto).
NIVEIS = {"INFO": logging.INFO, "CORREÇÃO": logging.INFO, "WARNING": logging.WARNING,
          "AVISO": logging.WARNING, "ERRO": logging.ERROR}

log = log_execucao.obter("main")


def _corrigir_excel(auxiliar, correcoes, mes, avisos):
    """Grava as correções na cópia da planilha. Falha aqui nunca impede o relatório."""
    if not correcoes:
        return
    try:
        destino, edicoes, puladas = corrigir_planilha(auxiliar["arquivo_producao"], auxiliar["meta"], correcoes, mes)
    except Exception as e:  # noqa: BLE001 - qualquer problema vira aviso; o Word não depende disso
        log.warning(f"[EXCEL][WARNING] Planilha NÃO corrigida: {e}")
        avisos["PENDENCIA"].append(f"Cópia corrigida da planilha de Produção não foi gerada: {e}")
        return
    for ed in edicoes:
        if ed["mesma"]:
            log.info(f"[EXCEL] {ed['celula']}: fórmula {ed['atual']} já correta; valor atualizado para {ed['valor']:.6g}.")
        else:
            log.info(f"[EXCEL] {ed['celula']}: {ed['atual']} -> {ed['formula']} (= {ed['valor']:.6g})")
    for p in puladas:
        log.warning(f"[EXCEL][WARNING] {p['celula']} ({p['chave']}) NÃO corrigida: {p['motivo']}")
        avisos["PENDENCIA"].append(f"Planilha: corrigir à mão {p['celula']} ({p['chave']}) — {p['motivo']}")
    if destino:
        log.info(f"[EXCEL] Planilha de Produção recalculada e salva com correções: {destino.name}")


def _ler_periodo(ap, args):
    try:
        periodo = resolver_periodo(args.mes, args.ano)
    except ValueError as e:
        ap.error(str(e))
    hoje = date.today()
    if (periodo.ano, periodo.numero) > (hoje.year, hoje.month):
        ap.error(f"{periodo.mes}/{periodo.ano} ainda não aconteceu (hoje é {hoje:%d/%m/%Y}).")
    return periodo


def _executar(args, periodo):
    mes, ano = periodo.mes, periodo.ano
    log.info(f"Referência: {mes}/{ano} ({periodo.n_dias} dias){' — modo estrito' if args.estrito else ''}")
    try:
        secoes, auxiliar = extrair_por_fonte(mes, ano, args.pasta_entrada)
    except (FileNotFoundError, ValueError) as e:
        log.error(f"Erro na extração (a planilha de Produção é indispensável): {e}")
        sys.exit(SAIDA_FALHA)

    avisos = {categoria: [] for categoria, _ in SECOES_AVISOS}
    hoje = date.today()
    if (ano, periodo.numero) == (hoje.year, hoje.month):
        avisos["CONFERIR"].append(f"{mes}/{ano} ainda não terminou: os dados do mês estão incompletos.")
    for fonte in auxiliar["pendencias"]:
        avisos["FONTE"].append(f"{fonte}: arquivo do mês não encontrado em {args.pasta_entrada.name}/; "
                               "as variáveis desta fonte estão como '[Sem Dados]'.")

    dados, registro, correcoes = consolidar_e_sanitizar(secoes, auxiliar["series"])
    log.info(f"Total de variáveis consolidadas: {len(dados)}")
    for nivel, msg in registro:
        log.log(NIVEIS.get(nivel, logging.INFO), f"[{nivel}] {msg}")
        if nivel == "WARNING":  # valor inválido que a sanitização não conseguiu recalcular
            avisos["CRITICO"].append(msg)

    _corrigir_excel(auxiliar, correcoes, mes, avisos)

    extras = {"mes": mes, "ano": ano, "n_dias": periodo.n_dias}
    secoes_finais = {titulo: {k: dados[k] for k in valores} for titulo, valores in secoes.items()}
    # Gabarito = exatamente o que vai ao Word. Com outra pasta de saída (ex.: o exemplo), vai junto dela.
    mapa = MAPA if args.pasta_saida == PASTA_SAIDA else args.pasta_saida / MAPA.name
    exportar_mapa(secoes_finais, mapa, extras=extras)

    if not TEMPLATE.is_file():
        log.warning(f"[WARNING] {TEMPLATE.name} não existe; criando modelo de exemplo (rode scripts/montar_template.py).")
        criar_template_exemplo(TEMPLATE, dados)

    problemas = validar(dados, auxiliar["series"], auxiliar["titulo_aba"], mes, ano,
                        dias_modelo=dias_no_modelo(TEMPLATE))
    for nivel, msg in problemas:
        log.log(NIVEIS.get(nivel, logging.WARNING), f"[{nivel}] {msg}")
        avisos["CRITICO" if nivel == "ERRO" else "CONFERIR"].append(msg)

    if args.estrito and any(avisos[c] for c in BLOQUEIAM_NO_ESTRITO):
        log.error("\n[ESTRITO] Relatório NÃO gerado. Resolva os itens abaixo e rode de novo:")
        for categoria in BLOQUEIAM_NO_ESTRITO:
            for msg in avisos[categoria]:
                log.error(f"  - [{categoria}] {msg}")
        sys.exit(SAIDA_BLOQUEADA)

    saida = args.pasta_saida / f"Relatorio_{mes.capitalize()}_{ano}_Final.docx"
    arquivo, faltantes, com_pagina = gerar_relatorio_word(
        dados, TEMPLATE, saida, extras=extras, series=auxiliar["series"], avisos=avisos
    )
    log.info(f"Relatório salvo em: {arquivo}")
    if faltantes:
        log.warning(f"[WARNING] {len(faltantes)} variável(is) do modelo sem valor ('[Sem Dados]' no Word): {faltantes}")
    if auxiliar["pendencias"]:
        log.warning(f"[WARNING] Fontes sem dados: {', '.join(auxiliar['pendencias'])}.")
    if com_pagina:
        log.warning("[WARNING] O Word abre com a página 'AVISOS CRÍTICOS E PENDÊNCIAS'; remova-a depois de resolver.")
    log.info("Campos manuais: busque '[PREENCHER' (amarelo) no Word.")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mes", help="mês de referência: setembro, set, 9 — ou com o ano: set/2026, 09-2026")
    ap.add_argument("ano", nargs="?", help="ano com 4 ou 2 dígitos (padrão: ano atual)")
    ap.add_argument("--estrito", action="store_true",
                    help="não gera o Word se houver erro não corrigível ou fonte ausente (código de saída 3)")
    ap.add_argument("--pasta-entrada", type=Path, default=PASTA_ENTRADA, metavar="PASTA",
                    help="onde estão as planilhas, o PDF e o ODT (padrão: dados_entrada/)")
    ap.add_argument("--pasta-saida", type=Path, default=PASTA_SAIDA, metavar="PASTA",
                    help="onde gravar o relatório Word e o log de auditoria (padrão: relatorios_gerados/)")
    args = ap.parse_args(argv)
    periodo = _ler_periodo(ap, args)

    caminho_log = log_execucao.configurar(args.pasta_saida, periodo)
    log.debug(f"Usuário: {getpass.getuser()} | máquina: {platform.node()} | Python {platform.python_version()}")
    log.debug(f"Comando: {' '.join([Path(sys.argv[0]).name] + (argv if argv is not None else sys.argv[1:]))}")
    codigo = 0
    try:
        _executar(args, periodo)
    except SystemExit as e:
        codigo = e.code if isinstance(e.code, int) else SAIDA_FALHA
        raise
    except Exception:
        codigo = SAIDA_FALHA
        log.exception("Falha inesperada. O detalhe técnico abaixo também está no log de auditoria.")
        raise SystemExit(SAIDA_FALHA) from None
    finally:
        log.debug(f"Fim da execução (código de saída {codigo}).")
        log.info(f"Log de auditoria: {caminho_log}")
        log_execucao.encerrar()


if __name__ == "__main__":
    main()
