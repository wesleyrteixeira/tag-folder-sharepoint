"""
scripts/listar_cnpjs.py
-----------------------
Consulta e exibe os CNPJs das pastas do 1º nível da biblioteca.

Uso:
    python scripts/listar_cnpjs.py
"""

import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.sharepoint import conectar, listar_cnpjs


def main():
    ctx = conectar()
    df  = listar_cnpjs(ctx)

    print(df.to_string(index=False))
    print(f"\nTotal        : {len(df)} pastas")
    print(f"Com CNPJ     : {(df['CNPJ'] != '').sum()}")
    print(f"Sem CNPJ     : {(df['CNPJ'] == '').sum()}")


if __name__ == "__main__":
    main()
