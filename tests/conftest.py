"""Configuração comum dos testes. Todos os dados são sintéticos: nenhum arquivo de dados_entrada/ é lido."""

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "scripts"))
sys.path.insert(0, str(RAIZ))  # main.py, para o teste ponta a ponta
