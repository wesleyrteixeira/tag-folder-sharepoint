"""
tools/gravar_cnpj_pasta.py
---------------------------
Grava o CNPJ em UMA pasta especifica definida via PASTA_URL no .env.
Usado para validar autenticacao e gravacao antes de rodar em massa.

Uso:
    python tools/gravar_cnpj_pasta.py

.env esperado:
    PASTA_URL=/sites/seu-site/Documentos Compartilhados/Nome da Pasta
"""

import sys
import os

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from src.sharepoint import conectar, garantir_coluna
from src.config import LIBRARY, COLUMN_NAME, PASTA_URL


CNPJ_TESTE = "01234567891234"


def gravar_cnpj_pasta(ctx):
    try:
        folder    = ctx.web.get_folder_by_server_relative_url(PASTA_URL)
        list_item = folder.list_item_all_fields
        ctx.load(list_item)
        ctx.execute_query()
    except Exception as e:
        print(f"Pasta nao encontrada: '{PASTA_URL}'")
        print(f"Detalhe: {e}")
        return

    item_id = list_item.properties.get("Id")
    nome    = list_item.properties.get("FileLeafRef")
    print(f"Pasta: {nome} (ID: {item_id})")

    lista = ctx.web.lists.get_by_title(LIBRARY)
    item  = lista.get_item_by_id(item_id)
    item.set_property(COLUMN_NAME, CNPJ_TESTE)
    item.update()
    ctx.execute_query()
    print(f"CNPJ gravado: {CNPJ_TESTE}")

    item_check = lista.get_item_by_id(item_id)
    ctx.load(item_check, [COLUMN_NAME, "FileLeafRef"])
    ctx.execute_query()
    print(f"Confirmacao: {item_check.properties.get(COLUMN_NAME)}")


if __name__ == "__main__":
    ctx = conectar()
    garantir_coluna(ctx)
    gravar_cnpj_pasta(ctx)
