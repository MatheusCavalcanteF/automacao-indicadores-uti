"""Localiza arquivos de entrada por extensão, termo base e mês/ano no nome (mtime só como fallback)."""

import re
import tempfile
import unicodedata
from pathlib import Path

# Nome completo (sem acento) -> apelidos aceitos no nome do arquivo.
MESES = {
    "JANEIRO": ["JANEIRO", "JAN", "01"],
    "FEVEREIRO": ["FEVEREIRO", "FEV", "02"],
    "MARCO": ["MARCO", "MAR", "03"],
    "ABRIL": ["ABRIL", "ABR", "04"],
    "MAIO": ["MAIO", "MAI", "05"],
    "JUNHO": ["JUNHO", "JUN", "06"],
    "JULHO": ["JULHO", "JUL", "07"],
    "AGOSTO": ["AGOSTO", "AGO", "08"],
    "SETEMBRO": ["SETEMBRO", "SET", "09"],
    "OUTUBRO": ["OUTUBRO", "OUT", "10"],
    "NOVEMBRO": ["NOVEMBRO", "NOV", "11"],
    "DEZEMBRO": ["DEZEMBRO", "DEZ", "12"],
}


class ArquivoAmbiguoError(Exception):
    """Arquivo só pôde ser escolhido por data do sistema, em modo estrito."""


def _sem_acento(texto):
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode()


def _tokens(nome):
    """'Fisio PED AGOSTO26 (2).pdf' -> ['FISIO','PED','AGOSTO','26','2','PDF']."""
    return re.findall(r"[A-Z]+|\d+", _sem_acento(nome).upper())


def resolver_mes(mes):
    """Aceita 'SETEMBRO', 'set', 'Setembro' ou '9'/'09' e devolve a chave de MESES."""
    alvo = _sem_acento(str(mes)).upper().strip()
    for chave, apelidos in MESES.items():
        if alvo == chave or alvo in apelidos or (alvo.isdigit() and int(alvo) == int(apelidos[2])):
            return chave
    raise ValueError(f"Mês inválido: {mes!r}")


def _anos(ano):
    ano = str(ano)
    return {ano, ano[-2:]}


def _meses_no_nome(tokens, anos):
    """Meses citados no nome. Número de mês ('09') só vale colado a um ano ('09-2026'),
    para não confundir com numeração de arquivo ('09 CENSO INPUT')."""
    achados = set()
    for i, tok in enumerate(tokens):
        for chave, apelidos in MESES.items():
            if tok in apelidos[:2]:
                achados.add(chave)
            elif tok == apelidos[2]:
                vizinhos = tokens[max(i - 1, 0):i] + tokens[i + 1:i + 2]
                if any(v in anos or len(v) == 4 and v.isdigit() for v in vizinhos):
                    achados.add(chave)
    return achados


def localizar_arquivo(pasta, extensao, termos, mes, ano, descricao=None, estrito=False):
    """Escolhe o arquivo `extensao` com algum de `termos` no nome, priorizando o do mês/ano.

    1. Candidatos: extensão + termo base (sem acento/caixa); ignora temporários do Office.
    2. Dentre eles, vence quem cita o mês no nome (+1 ponto extra se citar também o ano).
       Só o ano não basta: 'AGOSTO26' tem '26' mas é do mês errado.
    3. Sem match de mês: fallback pelo mais recente (mtime), com aviso forte; com
       `estrito=True` levanta ArquivoAmbiguoError em vez de seguir.
    """
    pasta = Path(pasta)
    if not pasta.is_dir():
        raise FileNotFoundError(f"Pasta de entrada não encontrada: {pasta}")

    mes_chave = resolver_mes(mes)
    anos = _anos(ano)
    termos_norm = [_sem_acento(t).lower() for t in termos]
    rotulo = descricao or "/".join(termos)

    candidatos = [
        f for f in pasta.iterdir()
        if f.is_file()
        and f.suffix.lower() == extensao.lower()
        and not f.name.startswith(("~$", ".~lock"))
        and not f.stem.endswith("_Corrigida")  # saída do write-back; a fonte é sempre o original
        and any(t in _sem_acento(f.name).lower() for t in termos_norm)
    ]
    if not candidatos:
        raise FileNotFoundError(
            f"Nenhum {extensao} com {' ou '.join(repr(t) for t in termos)} no nome em {pasta} ({rotulo})"
        )

    info = {}
    for f in candidatos:
        toks = _tokens(f.stem)
        meses = _meses_no_nome(toks, anos)
        info[f] = (mes_chave in meses, bool(anos & set(toks)), meses)

    com_mes = [f for f in candidatos if info[f][0]]
    if com_mes:
        # Mês+ano > só mês; empate resolvido por mtime, apenas para ser determinístico.
        escolhido = max(com_mes, key=lambda f: (info[f][1], f.stat().st_mtime))
        if len(com_mes) > 1:
            print(f"[aviso] {len(com_mes)} arquivos de {mes_chave} para {rotulo}; usando '{escolhido.name}'.")
        return escolhido

    escolhido = max(candidatos, key=lambda f: f.stat().st_mtime)
    outros = sorted(info[escolhido][2])
    if len(candidatos) == 1 and not outros:
        print(f"[info] {rotulo}: '{escolhido.name}' não cita mês no nome (arquivo único, provavelmente anual).")
    else:
        detalhe = f" O nome indica {', '.join(outros)}." if outros else ""
        if estrito:
            raise ArquivoAmbiguoError(
                f"{rotulo}: nenhum arquivo de {mes_chave} em {pasta}. O mais recente seria "
                f"'{escolhido.name}'.{detalhe}."
            )
        print(f"AVISO: Arquivo {escolhido.name} selecionado por data do sistema, "
              f"confira se o mês está correto (esperado: {mes_chave}).{detalhe}")
    return escolhido


def _testar():
    """Cenários sintéticos em pasta temporária."""
    import os
    import time

    with tempfile.TemporaryDirectory() as t:
        pasta = Path(t)

        def criar(nome, idade):
            f = pasta / nome
            f.write_text("x")
            os.utime(f, (time.time() - idade, time.time() - idade))

        criar("Relatório Fisio PED AGOSTO26.pdf", 0)       # mais recente, mês errado
        criar("Relatorio Fisio PED SET 2026.pdf", 500)     # mês certo, mais antigo
        criar("~$ Fisio SET.pdf", 0)                       # temporário: ignorado
        criar("Fisio 09-2026 final.pdf", 900)
        print("1)", localizar_arquivo(pasta, ".pdf", ["fisio"], "SETEMBRO", 2026, "fisio").name)
        print("2)", localizar_arquivo(pasta, ".pdf", ["fisio"], "OUTUBRO", 2026, "fisio").name)
        criar("09 CENSO INPUT AGOSTO 2026.xlsx", 0)
        print("3)", localizar_arquivo(pasta, ".xlsx", ["censo"], "SETEMBRO", 2026, "censo").name)
        criar("PRODUCAO DIARIA 2026.xlsx", 0)
        print("4)", localizar_arquivo(pasta, ".xlsx", ["producao"], "SETEMBRO", 2026, "produção").name)


if __name__ == "__main__":
    _testar()
    print("--- pasta real ---")
    pasta = Path(__file__).resolve().parent.parent / "dados_entrada"
    for ext, termos in ((".pdf", ["fisio"]), (".odt", ["nascidos", "testes"])):
        print(localizar_arquivo(pasta, ext, termos, "SETEMBRO", 2026).name)
