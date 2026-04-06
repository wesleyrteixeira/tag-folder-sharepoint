"""
scripts/listar_bibliotecas_sp.py
---------------------------------
Lista todas as listas e bibliotecas do site SharePoint.
Util para descobrir o nome exato da biblioteca antes de rodar outros scripts.

Uso:
    python scripts/listar_bibliotecas_sp.py
"""

import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.sharepoint import conectar


def main():
    ctx   = conectar()
    listas = ctx.web.lists.get().execute_query()

    print(f"\n{'='*60}")
    print(f"{'TITULO':<40} {'TEMPLATE':>10}")
    print(f"{'='*60}")
    for lista in listas:
        titulo   = lista.properties.get("Title", "")
        template = lista.properties.get("BaseTemplate", "")
        print(f"{titulo:<40} {template:>10}")


if __name__ == "__main__":
    main()
