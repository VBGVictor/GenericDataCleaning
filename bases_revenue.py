import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os
import re

"""
    Esta aplicação realiza a leitura, limpeza e classificação dos dados de um arquivo Excel em XLSM, onde,
    será retornado um dicionario de DataFrames de acordo com as tabelas que estiverem no arquivo (Caso estejam
    todas as tabelas em uma unica planilha).
"""

def clean_dataframe(df):
    """Limpa um DataFrame removendo linhas e colunas vazias e tratando os dados."""
    df.dropna(how='all', inplace=True)  
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col].fillna(0, inplace=True)
        else:
            df[col].fillna('', inplace=True)
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].apply(lambda x: re.sub(r'[^\w\s]', '', x))
    
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_numeric(df[col], errors='coerce')
        elif pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].str.strip()
    
    return df

def remove_empty_columns(df):
    """Remove apenas as colunas que estão completamente vazias (todos os valores são NaN ou strings vazias)."""
    if df is None or df.empty:
        return df
    df_clean = df.dropna(axis=1, how='all')
    return df_clean

def read_and_clean_xlsm(filepath):
    """Lê e limpa um arquivo XLSM, retornando um dicionário de DataFrames limpos junto aos seus nomes de identificação a depender da linha de cabeçalho."""
    try:
        xls = pd.ExcelFile(filepath, engine='openpyxl')
        dataframes = {}
        sheets = {
            "Produto 1": {"header": 1},     #Altere aqui
            "Produto 2": {"header": 2},     #Altere aqui
            "Produto 3": {"header": 0}      #Altere aqui
        }
        
        for sheet_name, options in sheets.items():
            if sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=options["header"])
                df = clean_dataframe(df)
                df = remove_empty_columns(df)
                dataframes[sheet_name] = df
            else:
                messagebox.showwarning("Aviso", f"A aba '{sheet_name}' não foi encontrada no arquivo.")
        return dataframes
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo não encontrado: {filepath}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
    return df

def select_file_and_process():
    """Abre um diálogo para selecionar um arquivo XLSM, lê os dados e retorna o caminho e o DataFrame processado."""
    root = tk.Tk()
    root.withdraw()  
    file_path = filedialog.askopenfilename(
        title="Selecione um arquivo XLSM",
        filetypes=[("Arquivos XLSM", "*.xlsm")],
        initialdir=os.path.expanduser("~\Downloads"))       #Altere aqui destino de uma pasta que o arquivo sempre estará
    root.destroy()  
    if file_path:
        try:
            df = pd.read_excel(file_path, engine='openpyxl') 
            df = read_and_clean_xlsm(file_path)  
            return file_path, df  
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
            return None, None
    else:
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
        return None, None


if __name__ == "__main__":
    file_path, dfs = select_file_and_process()
    if dfs:
        print("Arquivo processado com sucesso:", file_path)
        for sheet, df in dfs.items():
            print(f"\nPlanilha: {sheet}")
            print(df.head())  #Teste