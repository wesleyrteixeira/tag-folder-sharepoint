import pandas as pd
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.http.request_options import RequestOptions
from office365.runtime.http.http_method import HttpMethod

from src.config import CLIENT_ID, CLIENT_SECRET, SITE_URL, LIBRARY, COLUMN_NAME, COLUMN_LABEL
from src.excel import normalizar


# ---------------------------------------------------------------------------
# Conexão
# ---------------------------------------------------------------------------

def conectar() -> ClientContext:
    """Autentica no SharePoint via app-only (ClientId + ClientSecret)."""
    credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
    ctx = ClientContext(SITE_URL).with_credentials(credentials)
    ctx.web.get().execute_query()
    print(f"Conectado: {ctx.web.properties['Title']}")
    return ctx


# ---------------------------------------------------------------------------
# Coluna CNPJ
# ---------------------------------------------------------------------------

def garantir_coluna(ctx: ClientContext):
    """Cria a coluna CNPJ na biblioteca se ainda não existir (oculta da UI)."""
    lista  = ctx.web.lists.get_by_title(LIBRARY)
    fields = lista.fields.get().execute_query()

    nomes = [f.properties.get("InternalName", "") for f in fields]
    if COLUMN_NAME in nomes:
        print(f"Coluna '{COLUMN_NAME}' ja existe.")
        return

    schema_xml = (
        f'<Field Type="Text" '
        f'DisplayName="{COLUMN_LABEL}" '
        f'Name="{COLUMN_NAME}" '
        f'Hidden="TRUE" '
        f'ShowInViewForms="FALSE" '
        f'ShowInEditForm="FALSE" '
        f'ShowInNewForm="FALSE" />'
    )
    lista.fields.create_field_as_xml(schema_xml).execute_query()
    print(f"Coluna '{COLUMN_NAME}' criada (oculta).")


# ---------------------------------------------------------------------------
# Helpers REST internos
# ---------------------------------------------------------------------------

def _get_root_url(ctx: ClientContext) -> str:
    """Retorna a ServerRelativeUrl da raiz da biblioteca."""
    lista = ctx.web.lists.get_by_title(LIBRARY)
    ctx.load(lista, ["RootFolder"])
    ctx.execute_query()
    return lista.root_folder.properties["ServerRelativeUrl"]


def _rest_get(ctx: ClientContext, url: str) -> dict:
    """Executa GET REST autenticado e retorna o payload JSON."""
    request = RequestOptions(url)
    request.set_header("Accept", "application/json;odata=verbose")
    response = ctx.pending_request().execute_request_direct(request)
    response.raise_for_status()
    return response.json().get("d", {})


def _buscar_pastas_nivel1(ctx: ClientContext) -> list:
    """Retorna lista de (ID, nome) das pastas do 1º nível da biblioteca."""
    root_url = _get_root_url(ctx)
    base = f"{SITE_URL.rstrip('/')}/_api/web/lists/GetByTitle('{LIBRARY}')/items"
    url  = f"{base}?$select=ID,FileLeafRef,FSObjType,FileDirRef&$top=5000"

    pastas = []
    while url:
        payload = _rest_get(ctx, url)
        for item in payload.get("results", []):
            if item.get("FSObjType") == 1 and item.get("FileDirRef") == root_url:
                pastas.append((item["ID"], item.get("FileLeafRef", "")))
        url = payload.get("__next")

    return pastas


# ---------------------------------------------------------------------------
# Leitura
# ---------------------------------------------------------------------------

def listar_cnpjs(ctx: ClientContext) -> pd.DataFrame:
    """Retorna DataFrame com colunas ['Pasta', 'CNPJ'] do 1º nível da biblioteca."""
    root_url = _get_root_url(ctx)
    base = f"{SITE_URL.rstrip('/')}/_api/web/lists/GetByTitle('{LIBRARY}')/items"
    url  = f"{base}?$select=FileLeafRef,{COLUMN_NAME},FSObjType,FileDirRef&$top=5000"

    registros = []
    while url:
        payload = _rest_get(ctx, url)
        for item in payload.get("results", []):
            if item.get("FSObjType") == 1 and item.get("FileDirRef") == root_url:
                registros.append({
                    "Pasta": item.get("FileLeafRef", ""),
                    "CNPJ":  item.get(COLUMN_NAME) or "",
                })
        url = payload.get("__next")

    return pd.DataFrame(registros)


# ---------------------------------------------------------------------------
# Gravação
# ---------------------------------------------------------------------------

def gravar_cnpjs(ctx: ClientContext, mapeamento: dict, limite: int = None):
    """Grava o CNPJ nas pastas do 1º nível da biblioteca.

    Args:
        mapeamento: dict {nome_normalizado: cnpj} gerado por carregar_excel()
        limite:     se informado, processa no máximo N pastas (útil para testes)
    """
    pastas = _buscar_pastas_nivel1(ctx)
    print(f"Pastas no 1o nivel: {len(pastas)}")

    if limite:
        pastas = pastas[:limite]
        print(f"Limitado a {limite} pastas para teste.")

    lista        = ctx.web.lists.get_by_title(LIBRARY)
    encontrados  = 0
    nao_mapeados = []

    for item_id, nome_pasta in pastas:
        chave = normalizar(nome_pasta)
        if chave in mapeamento:
            cnpj = mapeamento[chave]
            item = lista.get_item_by_id(item_id)
            item.set_property(COLUMN_NAME, cnpj)
            item.update()
            ctx.execute_query()
            print(f"  OK  {nome_pasta}  ->  {cnpj}")
            encontrados += 1
        else:
            nao_mapeados.append(nome_pasta)

    print(f"\n{'='*60}")
    print(f"CNPJs gravados : {encontrados}")
    print(f"Sem mapeamento : {len(nao_mapeados)}")
    if nao_mapeados:
        print("\nPastas sem CNPJ no Excel:")
        for p in nao_mapeados:
            print(f"  - {p}")
