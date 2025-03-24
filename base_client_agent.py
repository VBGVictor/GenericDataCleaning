import pandas as pd
import tkinter as tk
from tkinter import filedialog
import chardet
import re

"""
    Esta aplicação realiza leitura do arquivo CSV, faz limpeza basica e verificações entregando um DataFrame
    pronto para uso.
"""

def detect_encoding(filepath):
    """Detecta o encoding do arquivo para evitar problemas de leitura."""
    try:
        with open(filepath, 'rb') as f:
            rawdata = f.read(10000)  
            encoding = chardet.detect(rawdata)['encoding']
        return encoding if encoding else 'utf-8'  
    except Exception as e:
        print(f"Erro na detecção do encoding: {e}")
        return 'utf-8'

def clean_column_names(df):
    """Normaliza os nomes das colunas para evitar problemas no PostgreSQL."""
    new_columns = []
    for col in df.columns:
        col_clean = re.sub(r'\W+', '_', col) 
        if col_clean[0].isdigit():  
            col_clean = f"col_{col_clean}"
        new_columns.append(col_clean)
    df.columns = new_columns
    return df

def read_and_clean_csv(filepath):
    """Lê o arquivo CSV e faz a limpeza dos dados e dos nomes das colunas."""
    encoding = detect_encoding(filepath)
    try:
        df = pd.read_csv(filepath, encoding=encoding)
    except Exception:
        print("Erro com separador ',' - Tentando com ';'")
        try:
            df = pd.read_csv(filepath, sep=';', encoding=encoding)
        except Exception as e:
            print(f"Erro ao ler o arquivo: {e}")
            return None
    df.columns = [col.strip() for col in df.columns]
    df = clean_column_names(df)
    df = df.applymap(lambda x: str(x).strip() if isinstance(x, str) else x).fillna('')
    return df

def filepath_target():
    root = tk.Tk()
    root.withdraw()
    initialdir = "C:/Users/Victor Goveia/Downloads/"  #coloque o caminho do seu arquivo para ficar de mais facil acesso
    filepath = filedialog.askopenfilename(
        title="Selecione um arquivo CSV",
        filetypes=[("CSV files", "*.csv")],
        initialdir=initialdir
    )
    
    return filepath

if __name__ == "__main__":
    filepath = filepath_target()
    if filepath:
        df = read_and_clean_csv(filepath)
        if df is not None:
            print("\nDataFrame Limpo:")
            print(df)
        else:
            print("Erro ao processar o arquivo.")
    else:
        print("Nenhum arquivo selecionado.")
