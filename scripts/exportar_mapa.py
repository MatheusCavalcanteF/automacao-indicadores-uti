"""Gera mapa_variaveis.txt com todas as variáveis consolidadas, por seção, para marcar o Word."""

import sys
from pathlib import Path

import log_execucao
from consolidacao import RAIZ, extrair_por_fonte
from utils import periodo_da_cli


def exportar_mapa(secoes, caminho_saida, extras=None):
    """Grava '{{ variavel }} = [valor]' por linha, com cabeçalho por seção.

    Na produção, uma linha em branco separa os blocos de indicadores (prefixo antes de '__').
    """
    linhas = []
    if extras:
        linhas.append("--- GERAL ---")
        linhas += [f"{{{{ {k} }}}} = [{v}]" for k, v in extras.items()]
        linhas.append("")

    for titulo, dados in secoes.items():
        linhas.append(f"--- {titulo} ---")
        bloco_anterior, composta_anterior = None, False
        for chave, valor in dados.items():
            bloco, composta = chave.split("__")[0], "__" in chave
            # Linha em branco só entre blocos com sub-linhas; indicadores soltos ficam juntos.
            if bloco_anterior is not None and bloco != bloco_anterior and (composta or composta_anterior):
                linhas.append("")
            bloco_anterior, composta_anterior = bloco, composta
            linhas.append(f"{{{{ {chave} }}}} = [{valor}]")
        linhas.append("")

    total = sum(len(d) for d in secoes.values()) + len(extras or {})
    linhas.append(f"Total: {total} variáveis")
    Path(caminho_saida).write_text("\n".join(linhas) + "\n", encoding="utf-8")
    return total


if __name__ == "__main__":
    log_execucao.configurar()
    periodo = periodo_da_cli(sys.argv[1:], "python scripts/exportar_mapa.py MES [ANO]")
    secoes, _ = extrair_por_fonte(periodo.mes, periodo.ano)
    destino = RAIZ / "mapa_variaveis.txt"
    n = exportar_mapa(secoes, destino, extras={"mes": periodo.mes, "ano": periodo.ano, "n_dias": periodo.n_dias})
    print(f"{n} variáveis gravadas em {destino}")
