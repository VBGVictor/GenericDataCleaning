import pandas as pd
import tkinter as tk
from tkinter import filedialog

"""
    Esta aplicação realiza leitura do arquivo XLSX, faz limpeza basica e verificações entregando um DataFrame
    pronto para uso.
"""

def read_and_clean_xlsx(filepath):
    try:
        df = pd.read_excel(filepath)
    except FileNotFoundError:
        print(f"Erro: Arquivo não encontrado em {filepath}")
        return None
    except pd.errors.EmptyDataError:
        print(f"Erro: Arquivo em {filepath} está vazio")
        return None
    except Exception as e:
        print(f"Um erro inesperado ocorreu ao ler o arquivo: {e}")
        return None
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col].fillna(0, inplace=True)
        else:
            df[col].fillna('', inplace=True)
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_numeric(df[col], errors='coerce')
        elif pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].str.strip()

    return df

def filepath_target():
    root = tk.Tk()
    root.withdraw()
    initialdir = "C:/Users/Victor Goveia/Downloads/"
    filepath = filedialog.askopenfilename(
        title="Selecione um arquivo XLSX",
        filetypes=[("XLSX files", "*.xlsx")],
        initialdir=initialdir
    )
    return filepath

if __name__ == "__main__":
    file = filepath_target()
    
    if file:
        df = read_and_clean_xlsx(file)
        if df is not None:
            print("\nDataFrame do arquivo XLSX:" , df)
        else:
            print("Não foi possível ler ou limpar o arquivo XLSX.")
    else:
        print("Nenhum arquivo selecionado.")
