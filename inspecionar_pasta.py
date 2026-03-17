"""
inspecionar_pasta.py
---------------------
Lê e exibe todos os metadados da pasta para confirmar como o CNPJ está armazenado.
"""

import os
import json
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

load_dotenv()

CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

SITE_URL     = os.getenv("SITE_URL")
LIBRARY      = os.getenv("LIBRARY")
PASTA_URL    = os.getenv("PASTA_URL")

credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
ctx = ClientContext(SITE_URL).with_credentials(credentials)
ctx.web.get().execute_query()
print(f"✅ Conectado: {ctx.web.properties['Title']}\n")

# Passo 1: pega o ID da pasta pelo path
folder    = ctx.web.get_folder_by_server_relative_url(PASTA_URL)
list_item = folder.list_item_all_fields
ctx.load(list_item)
ctx.execute_query()

item_id = list_item.properties.get("Id")
print(f"📁 ID da pasta: {item_id}\n")

# Passo 2: busca o item com todos os campos (incluindo customizados)
lista = ctx.web.lists.get_by_title(LIBRARY)
item  = lista.get_item_by_id(item_id)
ctx.load(item)
ctx.execute_query()

print(f"{'='*60}")
print(f"TODOS OS CAMPOS DO ITEM")
print(f"{'='*60}")
print(json.dumps(item.properties, indent=2, ensure_ascii=False, default=str))

# Passo 3: leitura direta e explícita do CNPJ
item2 = lista.get_item_by_id(item_id)
ctx.load(item2, ["CNPJ", "FileLeafRef"])
ctx.execute_query()

print(f"\n{'='*60}")
print(f"LEITURA DIRETA DO CAMPO CNPJ")
print(f"{'='*60}")
print(f"Pasta : {item2.properties.get('FileLeafRef')}")
print(f"CNPJ  : {item2.properties.get('CNPJ')}")