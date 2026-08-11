# Lectura, escritura y sanitización XML para Excel


import pandas as pd


def sanitizar_dataframe_para_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Limpia caracteres de control XML e invisibles que corrompen archivos Excel en Linux/Cloud."""
    df_clean = df.copy()
    for col in df_clean.columns:
        if df_clean[col].dtype == 'object':
            df_clean[col] = (
                df_clean[col]
                .astype(str)
                .str.replace(r'[\x00-\x08\x0B-\x0C\x0E-\x1F]', '', regex=True)
                .str.replace(r'_x[0-9a-fA-F]{4}_', '', regex=True)
            )
            df_clean[col] = df_clean[col].replace('nan', '')
    return df_clean