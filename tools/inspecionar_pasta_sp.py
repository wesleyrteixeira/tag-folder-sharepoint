"""
tools/inspecionar_pasta_sp.py
------------------------------
Exibe todos os metadados de uma pasta especifica (PASTA_URL no .env).
Util para debug e para entender o schema completo do item.

Uso:
    python tools/inspecionar_pasta_sp.py

.env esperado:
    PASTA_URL=/sites/seu-site/Documentos Compartilhados/Nome da Pasta
"""

import sys
import os
import json

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.sharepoint import conectar
from src.config import LIBRARY, COLUMN_NAME, PASTA_URL


def inspecionar_pasta(ctx):
    folder    = ctx.web.get_folder_by_server_relative_url(PASTA_URL)
    list_item = folder.list_item_all_fields
    ctx.load(list_item)
    ctx.execute_query()

    item_id = list_item.properties.get("Id")
    print(f"ID da pasta: {item_id}\n")

    lista = ctx.web.lists.get_by_title(LIBRARY)

    # Todos os campos padrao
    item = lista.get_item_by_id(item_id)
    ctx.load(item)
    ctx.execute_query()

    print("=" * 60)
    print("TODOS OS CAMPOS DO ITEM")
    print("=" * 60)
    print(json.dumps(item.properties, indent=2, ensure_ascii=False, default=str))

    # Campo CNPJ explicito (hidden precisa de select)
    item2 = lista.get_item_by_id(item_id)
    ctx.load(item2, [COLUMN_NAME, "FileLeafRef"])
    ctx.execute_query()

    print(f"\n{'='*60}")
    print("CAMPO CNPJ (leitura explicita)")
    print("=" * 60)
    print(f"Pasta : {item2.properties.get('FileLeafRef')}")
    print(f"CNPJ  : {item2.properties.get(COLUMN_NAME)}")


if __name__ == "__main__":
    ctx = conectar()
    inspecionar_pasta(ctx)
