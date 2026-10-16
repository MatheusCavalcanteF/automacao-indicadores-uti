"""Gera dados de entrada 100% FICTÍCIOS para rodar o pipeline sem nenhum dado real.

Cria em exemplo/dados_entrada/ (ou na pasta indicada):
  - PLANILHA PRODUCAO DIARIA EXEMPLO <ANO>.xlsx: aba do mês no layout que extracao_producao lê
    (título acima de DATAS, uma coluna por dia, TOTAL e REALIZADO, blocos de indicador);
  - CENSO EXEMPLO <MES> <ANO>.xlsx: aba "Resumo" com cabeçalho em três linhas, como extracao_censo lê.
PDF da Fisioterapia e ODT de Testes ficam de fora: o pipeline segue sem eles e marca "[Sem Dados]".

Valores aleatórios plausíveis (seed fixa) e três inconsistências de propósito, para o exemplo mostrar
a sanitização e a cópia corrigida da planilha:
  1. taxa de mortalidade digitada em % (com a seed padrão: 3,03 em vez de 0,0303);
  2. TOTAL do paciente-dia do bloco de CVC sem o último dia (e a taxa de CVC calculada com ele);
  3. paciente-dia do Censo diferente do da Produção (Produção é a fonte da verdade).

Uso: python scripts/gerar_dados_exemplo.py [MES [ANO]] [--pasta PASTA]   (padrão: setembro 2026)
"""

import argparse
import random
import sys
from pathlib import Path

import openpyxl

from utils import resolver_periodo

RAIZ = Path(__file__).resolve().parent.parent
PASTA_PADRAO = RAIZ / "exemplo" / "dados_entrada"
PASTA_REAL = RAIZ / "dados_entrada"
SEED = 2026
LEITOS = 10
UNIDADE = "Hospital Exemplo"


def _series(n_dias, rng):
    """Séries diárias coerentes entre si (crônicos <= pacientes <= leitos, eventos raros)."""
    s = {}
    s["cronicos"] = [rng.randint(1, 2) for _ in range(n_dias)]
    s["pacientes"] = [rng.randint(c + 5, LEITOS) for c in s["cronicos"]]
    s["leitos"] = [LEITOS] * n_dias
    s["leitos_sem_cronicos"] = [LEITOS - c for c in s["cronicos"]]
    s["pacientes_sem_cronicos"] = [p - c for p, c in zip(s["pacientes"], s["cronicos"])]
    s["saidas"] = [rng.choice((0, 1, 1, 1, 2)) for _ in range(n_dias)]
    s["vm"] = [rng.randint(2, min(5, p)) for p in s["pacientes"]]
    s["cvc"] = [rng.randint(3, min(6, p)) for p in s["pacientes"]]
    s["cvp"] = [rng.randint(2, min(5, p)) for p in s["pacientes"]]

    def eventos(n):
        dias = [0] * n_dias
        for d in rng.sample(range(n_dias), n):
            dias[d] = 1
        return dias

    s["obitos"] = eventos(1)
    s["reinternacoes"] = eventos(1)
    s["flebites"] = eventos(1)
    s["quedas"] = [0] * n_dias
    s["nascidos"] = [rng.randint(0, 4) for _ in range(n_dias)]
    return s


def _linhas(s):
    """[(rótulo, série diária ou None, TOTAL, REALIZADO)]; None entre blocos = linha em branco."""
    def conta(rotulo, serie, total=None):
        return (rotulo, serie, sum(serie) if total is None else total, None)

    def bloco(rotulo, realizado):
        return (rotulo, None, "TOTAL", realizado)

    pac, pac_sem, saidas = sum(s["pacientes"]), sum(s["pacientes_sem_cronicos"]), sum(s["saidas"])
    leitos, leitos_sem = sum(s["leitos"]), sum(s["leitos_sem_cronicos"])
    cvc_den_quebrado = pac - s["pacientes"][-1]  # inconsistência 2: a "fórmula" esqueceu o último dia
    return [
        conta("LEITO DIA GLOBAL", s["leitos"]),
        conta("LEITO DIA SEM OS CRÔNICOS", s["leitos_sem_cronicos"]),
        conta("PACIENTE DIA GLOBAL", s["pacientes"]),
        conta("PACIENTE DIA CRÔNICOS", s["cronicos"]),
        conta("N. DE NASCIDOS VIVOS", s["nascidos"]),
        None,
        bloco("SAÍDAS HOSPITALARES", saidas),
        conta("N. DE SAÍDAS", s["saidas"]),
        None,
        bloco("TAXA DE OCUPAÇÃO GLOBAL", pac / leitos),
        conta("N. DE PACIENTES DIA", s["pacientes"]),
        conta("N. DE LEITO DIAS", s["leitos"]),
        None,
        bloco("TAXA DE OCUPAÇÃO SEM OS CRÔNICOS", pac_sem / leitos_sem),
        conta("N. DE PACIENTES DIA", s["pacientes_sem_cronicos"]),
        conta("N. DE LEITO DIAS", s["leitos_sem_cronicos"]),
        None,
        bloco("MÉDIA DE PERMANÊNCIA", pac / saidas),
        conta("N. DE PACIENTE DIA", s["pacientes"]),
        conta("N. DE SAÍDAS", s["saidas"]),
        None,
        bloco("MÉDIA DE PERMANÊNCIA SEM OS CRÔNICOS", pac_sem / saidas),
        conta("N. DE PACIENTE DIA", s["pacientes_sem_cronicos"]),
        conta("N. DE SAÍDAS", s["saidas"]),
        None,
        bloco("TAXA DE MORTALIDADE EM UTI", 100 * sum(s["obitos"]) / saidas),  # inconsistência 1: x100
        conta("N. DE ÓBITOS", s["obitos"]),
        conta("N. TOTAL DE SAÍDAS", s["saidas"]),
        None,
        bloco("TAXA DE REINTERNAÇÃO EM 24H", sum(s["reinternacoes"]) / saidas),
        conta("N. DE REINTERNAÇÃO", s["reinternacoes"]),
        conta("N. TOTAL DE SAÍDAS", s["saidas"]),
        None,
        bloco("TAXA DE UTILIZAÇÃO DE VENTILAÇÃO MECÂNICA", sum(s["vm"]) / pac),
        conta("N. DE PACIENTES DIA EM VM", s["vm"]),
        conta("N. TOTAL DE PACIENTE DIA", s["pacientes"]),
        None,
        bloco("TAXA DE UTILIZAÇÃO DE CATETER VENOSO CENTRAL (CVC)", sum(s["cvc"]) / cvc_den_quebrado),
        conta("N. DE DIA DE CATETER CENTRAL", s["cvc"]),
        conta("N. DE PACIENTE DIA", s["pacientes"], total=cvc_den_quebrado),
        None,
        bloco("INCIDÊNCIA DE FLEBITE", sum(s["flebites"]) / sum(s["cvp"])),
        conta("N. DE FLEBITES", s["flebites"]),
        conta("N. DE PACIENTE DIA COM CATETER VENOSO PERIFÉRICO", s["cvp"]),
        None,
        bloco("INCIDÊNCIA DE QUEDA DE PACIENTE", sum(s["quedas"]) / pac),
        conta("N. DE INCIDÊNCIA DE QUEDA DE PACIENTES", s["quedas"]),
        conta("N. DE PACIENTE DIA", s["pacientes"]),
    ]


def gerar_producao(caminho, periodo, s):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"UTI PEDIATRICA {periodo.mes}"
    ws.append([f"{UNIDADE} - UTI PEDIÁTRICA - INDICADORES DE PRODUÇÃO DIÁRIA - {periodo.mes}/{periodo.ano}"])
    ws.append(["DADOS FICTÍCIOS GERADOS POR scripts/gerar_dados_exemplo.py"])
    ws.append(["DATAS", *range(1, periodo.n_dias + 1), "TOTAL", "REALIZADO"])
    for linha in _linhas(s):
        if linha is None:
            ws.append([])
            continue
        rotulo, serie, total, realizado = linha
        ws.append([rotulo, *(serie or [None] * periodo.n_dias), total, realizado])
    wb.save(caminho)


def gerar_censo(caminho, periodo, s):
    pac, leitos, saidas = sum(s["pacientes"]), sum(s["leitos"]), sum(s["saidas"])
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumo"
    ws.append([f"CENSO HOSPITALAR - {UNIDADE} - {periodo.mes}/{periodo.ano} (FICTÍCIO)"])
    ws.append(["Especialidade", "Leitos", "Paciente", "Total Saídas"])
    ws.append([None, "Dia", "Dia", "Geral"])
    ws.append([None, None, None, None])
    ws.append(["UTI NEONATAL", 20 * periodo.n_dias, 17 * periodo.n_dias, 41])
    ws.append(["UTI PEDIÁTRICA", leitos, pac + 3, saidas])  # inconsistência 3: paciente-dia diverge
    ws.append(["PEDIATRIA", 25 * periodo.n_dias, 19 * periodo.n_dias, 96])
    wb.save(caminho)


def gerar(pasta, mes="setembro", ano=2026, seed=SEED):
    """Grava os arquivos fictícios em `pasta` e devolve a lista de caminhos."""
    pasta = Path(pasta)
    if pasta.resolve() == PASTA_REAL.resolve():
        raise ValueError(f"{pasta} é a pasta dos dados reais; use outra (padrão: {PASTA_PADRAO}).")
    periodo = resolver_periodo(mes, ano)
    s = _series(periodo.n_dias, random.Random(seed))
    pasta.mkdir(parents=True, exist_ok=True)
    producao = pasta / f"PLANILHA PRODUCAO DIARIA EXEMPLO {periodo.ano}.xlsx"
    censo = pasta / f"CENSO EXEMPLO {periodo.mes} {periodo.ano}.xlsx"
    gerar_producao(producao, periodo, s)
    gerar_censo(censo, periodo, s)
    return [producao, censo]


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("mes", nargs="?", default="setembro")
    ap.add_argument("ano", nargs="?", default="2026")
    ap.add_argument("--pasta", type=Path, default=PASTA_PADRAO, help=f"destino (padrão: {PASTA_PADRAO})")
    args = ap.parse_args(argv)
    try:
        arquivos = gerar(args.pasta, args.mes, args.ano)
    except ValueError as e:
        ap.error(str(e))
    for arquivo in arquivos:
        print("Gerado:", arquivo)


if __name__ == "__main__":
    sys.exit(main())
