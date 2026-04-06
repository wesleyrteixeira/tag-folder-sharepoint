import re
import pandas as pd

from src.config import EXCEL_PATH, COL_PASTA, COL_CNPJ


def normalizar(texto: str) -> str:
    """Remove espaços extras e converte para maiúsculas para comparação."""
    if not isinstance(texto, str):
        return ""
    return re.sub(r"\s+", " ", texto.strip().upper())


def limpar_cnpj(cnpj) -> str:
    """Mantém apenas dígitos do CNPJ e garante 14 caracteres com zeros à esquerda."""
    return re.sub(r"\D", "", str(cnpj)).zfill(14)


def carregar_excel() -> dict:
    """Lê o Excel de mapeamento e retorna {nome_normalizado: cnpj}."""
    df = pd.read_excel(EXCEL_PATH, dtype=str)

    for col in [COL_PASTA, COL_CNPJ]:
        if col not in df.columns:
            raise ValueError(
                f"Coluna '{col}' nao encontrada no Excel. "
                f"Disponiveis: {list(df.columns)}"
            )

    df = df[[COL_PASTA, COL_CNPJ]].dropna(subset=[COL_PASTA])
    mapeamento = {
        normalizar(row[COL_PASTA]): limpar_cnpj(row[COL_CNPJ])
        for _, row in df.iterrows()
        if pd.notna(row[COL_CNPJ])
    }
    print(f"Excel carregado: {len(mapeamento)} entradas.")
    return mapeamento
