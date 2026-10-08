"""Configuração comum dos testes. Todos os dados são sintéticos: nenhum arquivo de dados_entrada/ é lido."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
