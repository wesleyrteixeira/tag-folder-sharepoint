# tag-pastas-sharepoint

Python script que grava CNPJs como metadado oculto em pastas de uma biblioteca do SharePoint, usando autenticação via Azure AD e mapeamento por planilha Excel.

---

## Contexto

O SharePoint armazena arquivos e pastas em **bibliotecas de documentos**, que por baixo são listas do SharePoint. Cada item dessa lista tem um conjunto de **campos/colunas** — alguns nativos (nome, data, autor) e outros customizados.

O objetivo deste projeto é associar o CNPJ de cada empresa à sua respectiva pasta na biblioteca, de forma **invisível para os usuários na UI**, mas **consultável via API** em automações futuras.

---

## Como funciona

### 1. Campo customizado oculto
Via SharePoint REST API, é adicionada uma coluna do tipo `Text` na biblioteca com os atributos `Hidden="TRUE"` e `ShowInViewForms="FALSE"` — o campo existe no schema da lista mas é **invisível na UI** para os usuários.

### 2. Busca de pastas via REST + FileDirRef
Para buscar apenas as pastas do **1º nível** sem varrer toda a biblioteca, usamos a REST API com `$select=FileLeafRef,CNPJ,FSObjType,FileDirRef` e filtramos em Python por `FileDirRef == root_url`. Isso evita o List View Threshold e não traz subpastas.

### 3. Gravação do metadado
Com o `ID` do item em mãos, acessamos via `get_item_by_id`, setamos o valor do campo `CNPJ` e executamos o `update()` — que dispara um `PATCH` na REST API do SharePoint.

### 4. Leitura
Campos customizados `Hidden=TRUE` não vêm no payload padrão. A leitura exige um `select` explícito:
```python
ctx.load(item, ["CNPJ"])  # gera $select=CNPJ na query da API
```

---

## Stack

| Biblioteca | Uso |
|---|---|
| `Office365-REST-Python-Client` | Wrapper Python para a SharePoint REST API |
| `pandas` + `openpyxl` | Leitura do Excel com `dtype=str` para preservar zeros à esquerda em CNPJs |
| `python-dotenv` | Credenciais via `.env` (autenticação app-only com `ClientId` + `ClientSecret` no Azure AD) |

---

## Pré-requisitos

### Azure AD — App Registration
O app precisa ter as seguintes **Application Permissions** concedidas (não delegadas):

- `Sites.ReadWrite.All`

### Instalação
```bash
pip install -r requirements.txt
```

### Arquivo `.env`
Crie um arquivo `.env` na raiz do projeto baseado no `.env.example`:
```
TENANT_ID=seu-tenant-id
CLIENT_ID=seu-client-id
CLIENT_SECRET=seu-client-secret
SITE_URL=https://seu-tenant.sharepoint.com/sites/seu-site
LIBRARY=Documentos
PASTA_URL=/sites/seu-site/Documentos Compartilhados/Nome da Pasta
EXCEL_PATH=C:\caminho\para\planilha.xlsx
```

---

## Estrutura do projeto

```
tag-pastas-sharepoint/
│
├── src/                              # Módulos reutilizáveis (lógica de negócio)
│   ├── config.py                     # Variáveis de ambiente e constantes
│   ├── excel.py                      # Leitura do Excel: normalizar, limpar_cnpj, carregar_excel
│   └── sharepoint.py                 # API SharePoint: conectar, garantir_coluna, listar_cnpjs, gravar_cnpjs
│
├── scripts/                          # Entry points — execução do dia a dia
│   ├── gravar_cnpjs.py               # Grava CNPJs em todas as pastas do 1º nível
│   ├── listar_cnpjs.py               # Consulta e exibe CNPJs como DataFrame
│   └── listar_bibliotecas_sp.py      # Utilitário: lista todas as bibliotecas do site
│
├── tools/                            # Scripts de debug e validação pontual
│   ├── gravar_cnpj_pasta.py          # Grava CNPJ em uma única pasta (PASTA_URL no .env)
│   ├── ler_cnpj_pasta.py             # Lê o CNPJ gravado em uma pasta específica
│   └── inspecionar_pasta_sp.py       # Exibe todos os metadados de uma pasta
│
├── .env                              # Credenciais e configurações (não versionar)
├── .env.example                      # Template do .env
├── requirements.txt                  # Dependências do projeto
└── README.md
```

**Convenção de nomenclatura:** `<verbo>_<objeto>_<contexto>.py` em snake_case.
- `src/` — nunca executado diretamente, apenas importado
- `scripts/` — execução em produção/dia a dia
- `tools/` — apenas para debug e validação pontual

---

## Scripts

### `scripts/gravar_cnpjs.py`
Script principal. Lê o Excel de mapeamento e grava o CNPJ em todas as pastas do 1º nível da biblioteca.

```bash
# Todas as pastas
python scripts/gravar_cnpjs.py

# Testar com N pastas antes de rodar em massa
python scripts/gravar_cnpjs.py --limite 5
```

**Excel esperado:**

| Pasta de arquivos | CNPJ |
|---|---|
| EMPRESA ALPHA LTDA | 01234567891234 |
| EMPRESA BETA S.A | 98765432000111 |

Output ao final:
```
CNPJs gravados : 45
Sem mapeamento : 3

Pastas sem CNPJ no Excel:
  - 2025
  - 2026
  - Modelos
```

---

### `scripts/listar_cnpjs.py`
Consulta e exibe os CNPJs gravados em todas as pastas do 1º nível como DataFrame.

```bash
python scripts/listar_cnpjs.py
```

Output esperado:
```
                        Pasta            CNPJ
     EMPRESA ALPHA LTDA       01234567891234
     EMPRESA BETA S.A         98765432000111

Total        : 47 pastas
Com CNPJ     : 45
Sem CNPJ     : 2
```

---

### `scripts/listar_bibliotecas_sp.py`
Lista todas as listas e bibliotecas disponíveis no site SharePoint.
Útil para descobrir o **nome exato da biblioteca** antes de rodar os outros scripts.

```bash
python scripts/listar_bibliotecas_sp.py
```

---

### `tools/gravar_cnpj_pasta.py`
Grava um CNPJ de teste em **uma única pasta** definida via `PASTA_URL` no `.env`.
Usado para validar autenticação, criação da coluna e gravação antes de rodar em massa.

```bash
python tools/gravar_cnpj_pasta.py
```

---

### `tools/ler_cnpj_pasta.py`
Lê e exibe o CNPJ gravado em uma pasta específica (`PASTA_URL` no `.env`).

```bash
python tools/ler_cnpj_pasta.py
```

Output esperado:
```
Pasta : Nome da Pasta
CNPJ  : 12345678901234
```

---

### `tools/inspecionar_pasta_sp.py`
Exibe **todos os metadados** de uma pasta (campos nativos e customizados).
Útil para debug e para entender o schema completo do item.

```bash
python tools/inspecionar_pasta_sp.py
```

> **Nota:** campos `Hidden=TRUE` não aparecem no payload padrão. O script faz uma segunda chamada com `select` explícito para garantir a leitura do campo `CNPJ`.

---

## Observações

- A coluna `CNPJ` é criada automaticamente na primeira execução, caso não exista.
- O script é **idempotente** — rodar mais de uma vez sobrescreve o valor sem criar duplicatas.
- Apenas pastas do **1º nível** são processadas. Subpastas (como `2025`, `2026`) são ignoradas — isso é esperado.
- CNPJs com zero à esquerda (ex: `01234567891234`) são preservados via `zfill(14)`.
