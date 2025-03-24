import pandas as pd
import re
from bases_revenue import select_file_and_process
import warnings

"""
    Esta aplicação realiza a organização das tabelas que estiverem com o cabeçario desajustados, processando o nome e o arquivo em questão.
    Depois concatena os DataFrames que antes vieram em dicionario em apenas um DataFrame totalmente ajustado e com nomes de identificação 
    das tabelas que contiam na planilha (os sheet's).
"""

def fix_header(df, header_row_index=0, null_threshold=0.5):
    """"Esta função ajusta o cabeçario caso tenha tabelas na mesma planilha desorganizadas. """
    total_cols = df.shape[1]
    df.dropna(axis=1, how='all', inplace=True)
    for idx in range(header_row_index, len(df)):
        row = df.iloc[idx]
        if row.isnull().sum() < null_threshold * total_cols:
            new_header = row.tolist()
            df_fixed = df.iloc[idx+1:].copy()
            df_fixed.columns = new_header
            df_fixed.dropna(how='all', inplace=True)
            df_fixed.dropna(axis=1, how='all', inplace=True)
            df_fixed.reset_index(drop=True, inplace=True)
            return df_fixed
    print("Nenhuma linha de cabeçalho adequada encontrada; retornando DataFrame original.")
    return df.copy()

def Process_df_date_roas():
    """"Esta função lê e processa o nome do arquivo para utilizar algum trecho que sirva de identificação daos dados, exemplo: datas. """
    x = int(input("Digite o número de arquivos a serem processados: "))
    dfs_all = []
    for i in range(x):
        file_path, df__dict = select_file_and_process()
        if file_path and df__dict:
            match = re.search(r'XXXX(\d{8})', file_path)        #Altere o 'XXXX' para a sequencia do nome proximo a info desejada
            if match:
                date_str = match.group(1)
                try:
                    date_obj = pd.to_datetime(date_str, format='%Y%m%d').date()
                except ValueError:
                    print(f"Data inválida no nome do arquivo: {file_path}. Verifique o formato (XXXAAAAMMDD).") #Altere o 'XXXX' para a sequencia do nome proximo a info desejada
                    continue
                try:
                    if isinstance(df__dict, dict):
                        for sheet_name, df in df__dict.items():
                            df_sheet = pd.read_excel(file_path, engine='openpyxl', header=None, sheet_name=sheet_name)
                            df_sheet_fixed = fix_header(df_sheet, header_row_index=0)
                            df_sheet_fixed["Data"] = date_obj
                            df_sheet_fixed.reset_index(drop=True, inplace=True)
                            df_sheet_fixed["SheetName"] = sheet_name
                            dfs_all.append(df_sheet_fixed)
                    else:
                        df_sheet = pd.read_excel(file_path, engine='openpyxl', header=None)
                        df_sheet_fixed = fix_header(df_sheet, header_row_index=0)
                        df_sheet_fixed["Data"] = date_obj
                        df_sheet_fixed.reset_index(drop=True, inplace=True)
                        dfs_all.append(df_sheet_fixed)
                    print(f"Arquivo {i+1} processado e corrigido.")
                except Exception as e:
                    print(f"Erro ao processar o arquivo {file_path}: {e}")
            else:
                print(f"Data não encontrada no nome do arquivo: {file_path}. Verifique o formato (AAI_AAAAMMDD).")
        else:
            print(f"Nenhum arquivo foi selecionado para o arquivo {i+1}.")

    print("Todos os arquivos foram processados.")
    if dfs_all:
        dfs_all_concat = pd.concat([df.reset_index(drop=True) for df in dfs_all], ignore_index=True)
        return dfs_all_concat
    else:
        print("Nenhum DataFrame válido foi processado.")
        return None

def main():
    warnings.simplefilter(action='ignore', category=FutureWarning)
    pd.set_option('mode.chained_assignment', None)
    df__final = Process_df_date_roas()
    if df__final is not None:
        print("Processamento concluído.")
    return df__final

if __name__ == "__main__":
    main()
