"""
tools/ler_cnpj_pasta.py
------------------------
Le e exibe o CNPJ gravado em uma pasta especifica (PASTA_URL no .env).
Usado para confirmar que o metadado foi persistido corretamente.

Uso:
    python tools/ler_cnpj_pasta.py

.env esperado:
    PASTA_URL=/sites/seu-site/Documentos Compartilhados/Nome da Pasta
"""

import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.sharepoint import conectar
from src.config import LIBRARY, COLUMN_NAME, PASTA_URL


def ler_cnpj_pasta(ctx):
    folder    = ctx.web.get_folder_by_server_relative_url(PASTA_URL)
    list_item = folder.list_item_all_fields
    ctx.load(list_item)
    ctx.execute_query()
    item_id = list_item.properties.get("Id")

    lista = ctx.web.lists.get_by_title(LIBRARY)
    item  = lista.get_item_by_id(item_id)
    ctx.load(item, [COLUMN_NAME, "FileLeafRef"])
    ctx.execute_query()

    print(f"Pasta : {item.properties.get('FileLeafRef')}")
    print(f"CNPJ  : {item.properties.get(COLUMN_NAME)}")


if __name__ == "__main__":
    ctx = conectar()
    ler_cnpj_pasta(ctx)
