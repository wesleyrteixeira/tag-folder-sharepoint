import os
from dotenv import load_dotenv

load_dotenv()

# Azure AD
TENANT_ID     = os.getenv("TENANT_ID")
CLIENT_ID     = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# SharePoint
SITE_URL   = os.getenv("SITE_URL")
LIBRARY    = os.getenv("LIBRARY")
PASTA_URL  = os.getenv("PASTA_URL")

# Coluna customizada
COLUMN_NAME  = "CNPJ"
COLUMN_LABEL = "CNPJ"

# Excel
EXCEL_PATH = os.getenv("EXCEL_PATH")
COL_PASTA  = "Pasta de arquivos"
COL_CNPJ   = "CNPJ"
