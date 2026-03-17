"""
teste_gravar_cnpj.py
---------------------
Teste pontual: grava o CNPJ de UMA pasta específica no SharePoint.

Dependências:
    pip install Office365-REST-Python-Client python-dotenv

.env esperado:
    TENANT_ID=...
    CLIENT_ID=...
    CLIENT_SECRET=...
"""

import os
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

load_dotenv()

CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

SITE_URL      = os.getenv("SITE_URL")
LIBRARY       = os.getenv("LIBRARY")
COLUMN_NAME   = "CNPJ"

# --- Entrada de teste ---
CNPJ_TESTE  = "01234567891234"

# ---------------------------------------------------------------------------
# Conexão
# ---------------------------------------------------------------------------

def conectar() -> ClientContext:
    credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
    ctx = ClientContext(SITE_URL).with_credentials(credentials)
    ctx.web.get().execute_query()
    print(f"✅ Conectado: {ctx.web.properties['Title']}")
    return ctx

# ---------------------------------------------------------------------------
# Criar coluna CNPJ se não existir (oculta)
# ---------------------------------------------------------------------------

def garantir_coluna(ctx: ClientContext):
    lista = ctx.web.lists.get_by_title(LIBRARY)
    fields = lista.fields.get().execute_query()

    nomes = [f.properties.get("InternalName", "") for f in fields]
    if COLUMN_NAME in nomes:
        print(f"ℹ️  Coluna '{COLUMN_NAME}' já existe — pulando criação.")
        return

    schema_xml = (
        f'<Field Type="Text" '
        f'DisplayName="{COLUMN_NAME}" '
        f'Name="{COLUMN_NAME}" '
        f'Hidden="TRUE" '
        f'ShowInViewForms="FALSE" '
        f'ShowInEditForm="FALSE" '
        f'ShowInNewForm="FALSE" />'
    )
    lista.fields.create_field_as_xml(schema_xml).execute_query()
    print(f"✅ Coluna '{COLUMN_NAME}' criada (oculta).")

# ---------------------------------------------------------------------------
# Buscar a pasta e gravar o CNPJ
# ---------------------------------------------------------------------------

def gravar_cnpj_teste(ctx: ClientContext):
    # Busca a pasta diretamente pelo caminho — evita o List View Threshold
    pasta_url = os.getenv("PASTA_URL")

    try:
        folder = ctx.web.get_folder_by_server_relative_url(pasta_url)
        list_item = folder.list_item_all_fields
        ctx.load(list_item)
        ctx.execute_query()
    except Exception as e:
        print(f"❌ Pasta não encontrada: '{pasta_url}'")
        print(f"   Detalhe: {e}")
        return

    item_id = list_item.properties.get("Id")
    nome    = list_item.properties.get("FileLeafRef")
    print(f"📁 Pasta encontrada: {nome} (ID: {item_id})")

    # Grava o CNPJ via list item
    lista    = ctx.web.lists.get_by_title(LIBRARY)
    item     = lista.get_item_by_id(item_id)
    item.set_property(COLUMN_NAME, CNPJ_TESTE)
    item.update()
    ctx.execute_query()
    print(f"✅ CNPJ gravado: {CNPJ_TESTE}")

    # Lê de volta para confirmar
    item_check = lista.get_item_by_id(item_id)
    ctx.load(item_check, [COLUMN_NAME, "FileLeafRef"])
    ctx.execute_query()
    print(f"🔍 Confirmação leitura: {item_check.properties.get(COLUMN_NAME)}")

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    ctx = conectar()
    garantir_coluna(ctx)
    gravar_cnpj_teste(ctx)