"""
ler_cnpj.py
------------
Lê o CNPJ gravado na pasta 31 - RESIDENCIAL MODELO.
"""

import os
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

load_dotenv()

CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

SITE_URL  = os.getenv("SITE_URL")
LIBRARY   = os.getenv("LIBRARY")
PASTA_URL = os.getenv("PASTA_URL")

credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
ctx = ClientContext(SITE_URL).with_credentials(credentials)
ctx.web.get().execute_query()

# Pega o ID pelo path
folder    = ctx.web.get_folder_by_server_relative_url(PASTA_URL)
list_item = folder.list_item_all_fields
ctx.load(list_item)
ctx.execute_query()
item_id = list_item.properties.get("Id")

# Lê o CNPJ explicitamente
lista = ctx.web.lists.get_by_title(LIBRARY)
item  = lista.get_item_by_id(item_id)
ctx.load(item, ["CNPJ", "FileLeafRef"])
ctx.execute_query()

print(f"Pasta : {item.properties.get('FileLeafRef')}")
print(f"CNPJ  : {item.properties.get('CNPJ')}")
