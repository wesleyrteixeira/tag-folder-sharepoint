"""
gravar_cnpj_sharepoint.py
--------------------------
1. Cria a coluna CNPJ na biblioteca do SharePoint (oculta da view padrão)
2. Lê o Excel local com o mapeamento Pasta → CNPJ
3. Faz match com as pastas existentes no SharePoint
4. Grava o CNPJ no metadado de cada pasta

Dependências:
    pip install Office365-REST-Python-Client pandas openpyxl python-dotenv

.env esperado:
    TENANT_ID=...
    CLIENT_ID=...
    CLIENT_SECRET=...
"""

import os
import re
import pandas as pd
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential

# ---------------------------------------------------------------------------
# Configuração
# ---------------------------------------------------------------------------

load_dotenv()

TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

SITE_URL      = os.getenv("SITE_URL")
LIBRARY       = os.getenv("LIBRARY")
COLUMN_NAME   = "CNPJ"          # nome interno da coluna no SharePoint
COLUMN_LABEL  = "CNPJ"          # nome de exibição (não aparece na view padrão)

EXCEL_PATH    = os.getenv("EXCEL_PATH")
COL_PASTA     = "Pasta de arquivos"
COL_CNPJ      = "CNPJ"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def normalizar(texto: str) -> str:
    """Remove espaços extras e converte para maiúsculas para comparação."""
    if not isinstance(texto, str):
        return ""
    return re.sub(r"\s+", " ", texto.strip().upper())


def limpar_cnpj(cnpj) -> str:
    """Mantém apenas dígitos do CNPJ."""
    return re.sub(r"\D", "", str(cnpj))


# ---------------------------------------------------------------------------
# Conexão SharePoint
# ---------------------------------------------------------------------------

def conectar() -> ClientContext:
    credentials = ClientCredential(CLIENT_ID, CLIENT_SECRET)
    ctx = ClientContext(SITE_URL).with_credentials(credentials)
    ctx.web.get().execute_query()
    print(f"✅ Conectado: {ctx.web.properties['Title']}")
    return ctx


# ---------------------------------------------------------------------------
# Criar coluna CNPJ (se não existir) — oculta da view padrão
# ---------------------------------------------------------------------------

def garantir_coluna(ctx: ClientContext):
    lista = ctx.web.lists.get_by_title(LIBRARY)
    fields = lista.fields.get().execute_query()

    nomes = [f.properties.get("InternalName", "") for f in fields]
    if COLUMN_NAME in nomes:
        print(f"ℹ️  Coluna '{COLUMN_NAME}' já existe — pulando criação.")
        return

    # Cria via CAML/XML para poder definir Hidden e não aparecer na view
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
    print(f"✅ Coluna '{COLUMN_NAME}' criada (oculta).")


# ---------------------------------------------------------------------------
# Carregar Excel
# ---------------------------------------------------------------------------

def carregar_excel() -> dict:
    """Retorna dict {nome_normalizado: cnpj_limpo}"""
    df = pd.read_excel(EXCEL_PATH, dtype=str)

    # Valida colunas
    for col in [COL_PASTA, COL_CNPJ]:
        if col not in df.columns:
            raise ValueError(f"Coluna '{col}' não encontrada no Excel. Colunas disponíveis: {list(df.columns)}")

    df = df[[COL_PASTA, COL_CNPJ]].dropna(subset=[COL_PASTA])
    mapeamento = {
        normalizar(row[COL_PASTA]): limpar_cnpj(row[COL_CNPJ])
        for _, row in df.iterrows()
        if pd.notna(row[COL_CNPJ])
    }
    print(f"📄 Excel carregado: {len(mapeamento)} entradas.")
    return mapeamento


# ---------------------------------------------------------------------------
# Buscar pastas no SharePoint e gravar CNPJ
# ---------------------------------------------------------------------------

def gravar_cnpjs(ctx: ClientContext, mapeamento: dict):
    lista = ctx.web.lists.get_by_title(LIBRARY)

    # Busca apenas itens do tipo pasta (FSObjType = 1)
    from office365.sharepoint.caml.query import CamlQuery
    query = CamlQuery()
    query.ViewXml = """
    <View Scope="RecursiveAll">
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name="FSObjType" />
                    <Value Type="Integer">1</Value>
                </Eq>
            </Where>
        </Query>
        <RowLimit>5000</RowLimit>
    </View>
    """
    items = lista.get_items(query).execute_query()
    print(f"📁 Pastas encontradas no SharePoint: {len(items)}")

    encontrados   = 0
    nao_mapeados  = []

    for item in items:
        nome_pasta = item.properties.get("FileLeafRef", "")
        chave      = normalizar(nome_pasta)

        if chave in mapeamento:
            cnpj = mapeamento[chave]
            item.set_property(COLUMN_NAME, cnpj)
            item.update()
            ctx.execute_query()
            print(f"  ✅ {nome_pasta}  →  {cnpj}")
            encontrados += 1
        else:
            nao_mapeados.append(nome_pasta)

    print(f"\n{'='*60}")
    print(f"✅ CNPJs gravados : {encontrados}")
    print(f"⚠️  Sem mapeamento: {len(nao_mapeados)}")

    if nao_mapeados:
        print("\nPastas sem CNPJ no Excel:")
        for p in nao_mapeados:
            print(f"  - {p}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    ctx = conectar()
    garantir_coluna(ctx)
    mapeamento = carregar_excel()
    gravar_cnpjs(ctx, mapeamento)