"""
listar_bibliotecas.py
----------------------
Lista todas as listas/bibliotecas disponíveis no site para descobrir o nome correto.
"""

import os
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

load_dotenv()

CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_URL      = os.getenv("SITE_URL")

credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
ctx = ClientContext(SITE_URL).with_credentials(credentials)

listas = ctx.web.lists.get().execute_query()

print(f"\n{'='*60}")
print(f"{'TÍTULO':<40} {'TEMPLATE':>10}")
print(f"{'='*60}")
for lista in listas:
    titulo    = lista.properties.get("Title", "")
    template  = lista.properties.get("BaseTemplate", "")
    print(f"{titulo:<40} {template:>10}")