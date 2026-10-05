"""Sanitização dos dados consolidados: corrige fórmulas quebradas e concilia Censo x Produção.

Cada correção também é devolvida de forma estruturada, para o write-back na planilha:
  {"chave", "novo", "acao"} com acao = ("soma_linha",) | ("div", num, den) | ("sub", a, b) | ("ref", a).

Roda DEPOIS da extração e ANTES de validar/gerar o Word. Toda alteração é registrada:
  [INFO]      divergência esperada resolvida por regra de negócio (Produção é a fonte da verdade);
  [CORREÇÃO]  valor recalculado a partir das variáveis do próprio bloco;
  [WARNING]   valor inválido sem como recalcular: vira marcador de pendência (nunca um chute).
"""

import re

# Indicadores que são proporções (0–1). Só vale para a linha-título do bloco (sem "__").
PREFIXOS_PROPORCAO = (
    "taxa_", "incidencia_", "indice_", "adesao_", "pnm_", "atender_100", "fonoaudiologo_",
)
# (taxa, numerador, denominador): taxa = numerador / denominador. Todas conferidas contra relatórios reais.
RAZOES = [
    ("taxa_de_ocupacao_global", "taxa_de_ocupacao_global__n_de_pacientes_dia", "taxa_de_ocupacao_global__n_de_leito_dias"),
    ("taxa_de_ocupacao_sem_os_cronicos", "taxa_de_ocupacao_sem_os_cronicos__n_de_pacientes_dia", "taxa_de_ocupacao_sem_os_cronicos__n_de_leito_dias"),
    ("media_de_permanencia", "media_de_permanencia__n_de_paciente_dia", "media_de_permanencia__n_de_saidas"),
    ("media_de_permanencia_sem_os_cronicos", "media_de_permanencia_sem_os_cronicos__n_de_paciente_dia", "media_de_permanencia_sem_os_cronicos__n_de_saidas"),
    ("taxa_de_mortalidade_em_uti", "taxa_de_mortalidade_em_uti__n_de_obitos", "taxa_de_mortalidade_em_uti__n_total_de_saidas"),
    ("taxa_de_reinternacao_em_24h", "taxa_de_reinternacao_em_24h__n_de_reinternacao", "taxa_de_reinternacao_em_24h__n_total_de_saidas"),
    ("taxa_de_utilizacao_de_ventilacao_mecanica", "taxa_de_utilizacao_de_ventilacao_mecanica__n_de_pacientes_dia_em_vm", "taxa_de_utilizacao_de_ventilacao_mecanica__n_total_de_paciente_dia"),
    ("taxa_de_utilizacao_de_cateter_venoso_central_cvc", "taxa_de_utilizacao_de_cateter_venoso_central_cvc__n_de_dia_de_cateter_central", "taxa_de_utilizacao_de_cateter_venoso_central_cvc__n_de_paciente_dia"),
    ("incidencia_de_flebite", "incidencia_de_flebite__n_de_flebites", "incidencia_de_flebite__n_de_paciente_dia_com_cateter_venoso_periferico"),
    ("incidencia_de_perda_de_cateter_central_insercao_periferica", "incidencia_de_perda_de_cateter_central_insercao_periferica__n_de_perda_de_cateter_central_de_insercao_periferica", "incidencia_de_perda_de_cateter_central_insercao_periferica__n_de_paciente_dia_com_picc"),
]
# Linha da Produção que vale no lugar da do Censo.
CENSO_PRODUCAO = {
    "censo_leitos_dia": "leito_dia_global",
    "censo_paciente_dia": "paciente_dia_global",
    "censo_total_saidas_geral": "saidas_hospitalares",
}
# Adesão: o único numerador do bloco é dividido pelo paciente-dia global.
RECALCULO_ESPECIAL = {
    "adesao_as_metas_de_identificacao_do_paciente": (
        "adesao_as_metas_de_identificacao_do_paciente__utilizar_a_pulseira_de_identificacao_e_a_dupla_chegagem_de_medicamentos",
        "paciente_dia_global",
    ),
}

PENDENTE_INVALIDO = "[Valor Inválido na Planilha]"

_PACIENTE_DIA = re.compile(r"^n_(total_de_)?(de_)?pacientes?_dia$")
_LEITO_DIA = re.compile(r"^n_de_leito_dias?$")


def _num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _fmt(v):
    return f"{v:.4g}" if _num(v) else str(v)


def _conciliar_censo(dados, log):
    """Produção é a fonte da verdade: divergência vira [INFO] e o valor do Censo é substituído."""
    for censo, producao in CENSO_PRODUCAO.items():
        if _num(dados.get(censo)) and _num(dados.get(producao)) and dados[censo] != dados[producao]:
            log.append(("INFO", f"Censo x Produção divergem em {censo} (censo={_fmt(dados[censo])}, "
                                f"produção={_fmt(dados[producao])}); adotado o valor da Produção."))
            dados[censo] = dados[producao]


def _fonte_do_filho(chave, dados):
    """Valor autoritário de um denominador/numerador repetido dentro de um bloco, ou None."""
    bloco, _, nome = chave.partition("__")
    if not nome:
        return None
    sem_cronicos = "cronicos" in bloco
    if _PACIENTE_DIA.match(nome):
        base = dados.get("paciente_dia_global")
        cron = dados.get("paciente_dia_cronicos")
        if sem_cronicos:
            return base - cron if _num(base) and _num(cron) else None
        return base
    if _LEITO_DIA.match(nome):
        return dados.get("leito_dia_sem_os_cronicos" if sem_cronicos else "leito_dia_global")
    return None


def _acao_filho(chave, esperado, dados, series):
    """Como expressar a correção no Excel: soma da própria linha, se bate; senão referência à fonte."""
    dias = [v for v in (series or {}).get(chave, []) if v is not None]
    if dias and abs(sum(dias) - esperado) < 0.5:
        return ("soma_linha",)
    if "cronicos" in chave.partition("__")[0]:
        return ("sub", "paciente_dia_global", "paciente_dia_cronicos")
    return ("ref", "paciente_dia_global" if _PACIENTE_DIA.match(chave.partition("__")[2]) else
            "leito_dia_sem_os_cronicos" if "cronicos" in chave else "leito_dia_global")


def _corrigir_filhos(dados, log, correcoes, series):
    """'N. DE PACIENTE DIA' dentro dos blocos deve ser o paciente-dia global (ou sem crônicos)."""
    for chave in list(dados):
        esperado = _fonte_do_filho(chave, dados)
        if _num(esperado) and _num(dados[chave]) and abs(dados[chave] - esperado) > 0.5:
            log.append(("CORREÇÃO", f"{chave}: planilha trouxe {_fmt(dados[chave])}; "
                                    f"recalculado para {_fmt(esperado)} (total do próprio mês)."))
            correcoes.append({"chave": chave, "novo": esperado, "acao": _acao_filho(chave, esperado, dados, series)})
            dados[chave] = esperado


def _recalcular_razoes(dados, log, correcoes):
    for taxa, num, den in RAZOES:
        if not all(_num(dados.get(k)) for k in (taxa, num, den)) or not dados[den]:
            continue
        esperado = dados[num] / dados[den]
        if abs(esperado - dados[taxa]) > 0.005 * max(1, abs(esperado)):
            log.append(("CORREÇÃO", f"{taxa}: planilha trouxe {_fmt(dados[taxa])}; "
                                    f"recalculado {num.split('__')[1]} / {den.split('__')[1]} = {_fmt(esperado)}."))
            correcoes.append({"chave": taxa, "novo": esperado, "acao": ("div", num, den)})
            dados[taxa] = esperado


def _normalizar_proporcoes(dados, log, correcoes):
    """Proporção fora de 0–1: recalcula pelo bloco; sem como recalcular, marca pendente.

    Não divide por 100: 2,57 numa taxa que deveria ser 100% viraria 2,57%, outro número errado.
    """
    for chave in list(dados):
        valor = dados[chave]
        if "__" in chave or not chave.startswith(PREFIXOS_PROPORCAO) or not _num(valor) or 0 <= valor <= 1:
            continue
        if chave in RECALCULO_ESPECIAL:
            num_k, den_k = RECALCULO_ESPECIAL[chave]
        else:  # bloco com exatamente dois componentes: primeiro / segundo
            filhos = [k for k in dados if k.startswith(chave + "__")]
            num_k, den_k = filhos[:2] if len(filhos) == 2 else (None, None)
        num, den = dados.get(num_k), dados.get(den_k)
        novo = num / den if _num(num) and _num(den) and den else None
        if novo is not None and 0 <= novo <= 1:
            log.append(("CORREÇÃO", f"{chave}: planilha trouxe {_fmt(valor)} (fora de 0–100%); "
                                    f"recalculado {num_k.split('__')[-1]} / {den_k.split('__')[-1]} = {_fmt(novo)}."))
            correcoes.append({"chave": chave, "novo": novo, "acao": ("div", num_k, den_k)})
            dados[chave] = novo
        else:
            log.append(("WARNING", f"{chave} = {_fmt(valor)} fora de 0–100% e sem como recalcular; "
                                   f"marcado como pendente."))
            dados[chave] = PENDENTE_INVALIDO


def sanitizar(dados, series=None):
    """Devolve (dados corrigidos, log, correções estruturadas). Não altera o dicionário recebido."""
    dados, log, correcoes = dict(dados), [], []
    _conciliar_censo(dados, log)
    _corrigir_filhos(dados, log, correcoes, series)
    _recalcular_razoes(dados, log, correcoes)
    _normalizar_proporcoes(dados, log, correcoes)
    return dados, log, correcoes
