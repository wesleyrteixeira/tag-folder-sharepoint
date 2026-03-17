# tag-pastas-sharepoint

Projeto Python para gravar o **CNPJ de empresas como metadado oculto** nas pastas de uma biblioteca do SharePoint, sem alterar o nome das pastas.

---

## Contexto

O SharePoint armazena arquivos e pastas em **bibliotecas de documentos**, que por baixo são listas do SharePoint. Cada item dessa lista tem um conjunto de **campos/colunas** — alguns nativos (nome, data, autor) e outros customizados.

O objetivo deste projeto é associar o CNPJ de cada empresa à sua respectiva pasta na biblioteca, de forma **invisível para os usuários na UI**, mas **consultável via API** em automações futuras.

---

## O que foi feito

### 1. Campo customizado oculto
Via SharePoint REST API, foi adicionada uma coluna do tipo `Text` na biblioteca com os atributos `Hidden="TRUE"` e `ShowInViewForms="FALSE"` — o campo existe no schema da lista mas é **invisível na UI** para os usuários.

### 2. Navegação por path (evitando List View Threshold)
Em vez de fazer um `$filter` na lista inteira (que bate no limite de 5.000 itens do SharePoint), navegamos diretamente pelo **server-relative URL** de cada pasta com `get_folder_by_server_relative_url`. Isso retorna o `list_item_all_fields` da pasta, de onde extraímos o `Id` do item sem disparar o throttle.

### 3. Gravação do metadado
Com o `Id` em mãos, acessamos o item via `get_item_by_id`, setamos o valor do campo `CNPJ` e executamos o `update()` — que dispara um `PATCH` na REST API do SharePoint.

### 4. Leitura
Campos customizados não vêm no payload padrão. A leitura exige um `select` explícito:
```python
ctx.load(item, ["CNPJ"])  # gera $select=CNPJ na query da API
```

---

## Stack

| Biblioteca | Uso |
|---|---|
| `Office365-REST-Python-Client` | Wrapper Python para a SharePoint REST API |
| `pandas` + `openpyxl` | Leitura do Excel com tipagem forçada (`dtype=str`) para preservar zeros à esquerda em CNPJs |
| `python-dotenv` | Credenciais via `.env` (autenticação app-only com `ClientId` + `ClientSecret` no Azure AD) |

---

## Pré-requisitos

### Azure AD — App Registration
O app precisa ter as seguintes **Application Permissions** concedidas (não delegadas):

- `Sites.ReadWrite.All`

### Instalação
```bash
pip install Office365-REST-Python-Client pandas openpyxl python-dotenv
```

### Arquivo `.env`
Crie um arquivo `.env` na raiz do projeto (baseado no `.env.example`):
```
TENANT_ID=seu-tenant-id
CLIENT_ID=seu-client-id
CLIENT_SECRET=seu-client-secret
```

---

## Estrutura do projeto

```
tag-pastas-sharepoint/
├── .env                          # credenciais (não versionar)
├── .env.example                  # template do .env
├── listar_bibliotecas.py         # utilitário: lista todas as bibliotecas do site
├── teste_gravar_cnpj.py          # teste pontual com uma única pasta
├── ler_cnpj.py                   # lê e confirma o CNPJ gravado em uma pasta
├── inspecionar_pasta.py          # inspeciona todos os metadados de uma pasta
└── gravar_cnpj_sharepoint.py     # script principal: processa todas as pastas via Excel
```

---

## Scripts

### `listar_bibliotecas.py`
Lista todas as listas e bibliotecas disponíveis no site SharePoint.
Útil para descobrir o **nome exato da biblioteca** antes de rodar os outros scripts.

```bash
python listar_bibliotecas.py
```

---

### `teste_gravar_cnpj.py`
Teste pontual que grava o CNPJ em **uma única pasta** hardcoded.
Usado para validar a autenticação, a criação da coluna e a gravação antes de rodar em massa.

```bash
python teste_gravar_cnpj.py
```

Configurações no topo do arquivo:
```python
PASTA_TESTE = "31 - RESIDENCIAL MODELO"
CNPJ_TESTE  = "12345678901234"
```

---

### `ler_cnpj.py`
Lê e exibe o CNPJ gravado em uma pasta específica.
Confirma que o metadado foi persistido corretamente.

```bash
python ler_cnpj.py
```

Output esperado:
```
Pasta : 31 - RESIDENCIAL MODELO
CNPJ  : 12345678901234
```

---

### `inspecionar_pasta.py`
Exibe **todos os metadados** de uma pasta (campos nativos e customizados).
Útil para debug e para entender o schema completo do item.

```bash
python inspecionar_pasta.py
```

> **Nota:** campos customizados não aparecem no `list_item_all_fields` padrão.
> O script faz uma segunda chamada com `select` explícito para garantir a leitura do campo `CNPJ`.

---

### `gravar_cnpj_sharepoint.py`
Script principal. Lê o Excel de mapeamento e grava o CNPJ em todas as pastas da biblioteca.

```bash
python gravar_cnpj_sharepoint.py
```

**Excel esperado:**

| Pasta de arquivos | CNPJ |
|---|---|
| 31 - RESIDENCIAL MODELO | 01234567891234 |
| 32 - OUTRA EMPRESA LTDA | 98765432000111 |

> ⚠️ A coluna `CNPJ` no Excel deve estar formatada como **Texto** para preservar zeros à esquerda. O script aplica `dtype=str` na leitura e `zfill(14)` como proteção adicional.

Output ao final:
```
✅ CNPJs gravados : 45
⚠️  Sem mapeamento: 3

Pastas sem CNPJ no Excel:
  - 2025
  - 2026
  - Modelos
```

---

## Observações

- A coluna `CNPJ` é criada automaticamente na primeira execução, caso não exista.
- O script é **idempotente** — rodar mais de uma vez sobrescreve o valor sem criar duplicatas.
- Subpastas (como `2025`, `2026`) não terão CNPJ mapeado — isso é esperado.
- CNPJs com zero à esquerda (ex: `01234567891234`) são preservados corretamente via `zfill(14)`.
