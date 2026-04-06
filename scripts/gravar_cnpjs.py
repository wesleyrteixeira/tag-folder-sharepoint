"""
scripts/gravar_cnpjs.py
-----------------------
Grava o CNPJ nas pastas do 1º nível da biblioteca do SharePoint.

Uso:
    python scripts/gravar_cnpjs.py              # todas as pastas
    python scripts/gravar_cnpjs.py --limite 5   # teste com 5 pastas
"""

import argparse
import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.sharepoint import conectar, garantir_coluna, gravar_cnpjs
from src.excel import carregar_excel


def main():
    parser = argparse.ArgumentParser(description="Grava CNPJs nas pastas do SharePoint.")
    parser.add_argument("--limite", type=int, default=None, help="Limita o numero de pastas processadas (teste)")
    args = parser.parse_args()

    ctx = conectar()
    garantir_coluna(ctx)
    mapeamento = carregar_excel()
    gravar_cnpjs(ctx, mapeamento, limite=args.limite)


if __name__ == "__main__":
    main()
