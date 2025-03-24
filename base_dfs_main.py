import pandas as pd
import bases_date_revenue
import base_client_agent
from base_operations import read_and_clean_xlsx, filepath_target
import os

"""
    Esta aplicação realiza o precessamento inicial de Dataframes que serão utilizados, extratindo da parte de texto um codigo de produto especifico
    para ser utilizado como filtro e encontrar os valores de receita de cada produto de acordo com o dia (pois mudam a cada dia). Além disso, 
    cria-se um filtro apenas com os agentes e seus respectivos clientes e na execução final tambem salva o DataFrame principal com todos os agentes
    junto com o filtrado em um arquivo csv na area de trabalho do computador.
"""

def Process_files():
    file_xlsx = filepath_target()
    df_operations = read_and_clean_xlsx(file_xlsx)
    file_csv = base_client_agent.filepath_target()
    df_clients = base_client_agent.read_and_clean_csv(file_csv)
    df_revenue_combined = bases_date_revenue.main()
    return df_clients, df_operations, df_revenue_combined

def extract_valid_code(produto, valid_codes):
    if not isinstance(produto, str):
        return None
    produto = produto.upper().strip() 
    for code in valid_codes:
        code_str = str(code).upper().strip()  
        if not code_str:  
            continue
        if code_str in produto:  
            return code_str  
    return None

def convert_to_float(value):
    try:
        s = str(value).strip()
        s = s.replace('%', '').replace(',', '.')
        num = float(s)
    except ValueError:
        return None
    while num > 10:
        num /= 10
    if num > 6:
        num /= 1000
    elif num > 1:
        num /= 100
    return round(num, 4)

def handling_data():

    ### Segue um exemplo de ativação e uso das funções para utilizar a API e montar a base de dados da maneira desejada ###

    df_clients, df_operations, df_revenue_combined = Process_files()
    pd.set_option('display.float_format', '{:.4f}'.format) # Caso precise
    if 'CÓDIGO' not in df_revenue_combined.columns:
        print("Erro: A coluna 'CÓDIGO' não existe em df_revenue_combined!")
        return None

    valid_codes = df_revenue_combined['CÓDIGO'].unique()
    valid_codes = [code for code in valid_codes if str(code).strip() != ""]

    df_clients['CLIENTE'] = df_clients['CLIENTE'].str.strip().str.lower()
    df_operations['CLIENTE'] = df_operations['CLIENTE'].str.strip().str.lower()
    df_clients = df_clients.drop_duplicates(subset=['CLIENTE'])
    
    df_operations = df_operations[df_operations['MOVIMENTAÇÃO'] == 'CRÉDITO']
    df_operations = df_operations.merge(df_clients[['CLIENTE', 'AGENTE']], on='CLIENTE', how='left')
    df_operations['CÓDIGO_EXTRAIDO'] = df_operations['PRODUTOS'].apply(lambda x: extract_valid_code(x, valid_codes))
    df_operations['DATA DE PAGAMENTO'] = pd.to_datetime(df_operations['DATA DE PAGAMENTO'], errors='coerce')
    df_revenue_combined['DATA DO RENDIMENTO'] = pd.to_datetime(df_revenue_combined['DATA DO RENDIMENTO'], errors='coerce')
    df_operations['DATA DE PAGAMENTO'] = df_operations['DATA DE PAGAMENTO'].dt.date
    df_revenue_combined['DATA DO RENDIMENTO'] = df_revenue_combined['DATA DO RENDIMENTO'].dt.date

    df_operations = df_operations.merge(
        df_revenue_combined[['CÓDIGO', 'RECEITA ESTIMADA', 'DATA DO RENDIMENTO']],
        left_on=['CÓDIGO_EXTRAIDO', 'DATA DO RENDIMENTO'],
        right_on=['CÓDIGO', 'DATA DO RENDIMENTO'],
        how='outer'
    )

    df_operations['VALOR LÍQUIDO'] = pd.to_numeric(df_operations['VALOR LÍQUIDO'], errors='coerce')
    df_operations['RECEITA ESTIMADA'] = df_operations['RECEITA ESTIMADA'].apply(convert_to_float)
    df_operations = df_operations[df_operations['RECEITA ESTIMADA'].notna()]
    df_operations['RECEITA ESTIMADA'] = df_operations['RECEITA ESTIMADA'].round(4)
    df_operations['RECEITA LÍQUIDA'] = (((df_operations['RECEITA ESTIMADA'] * df_operations['VALOR LÍQUIDO']) *0,5)).astype('float64').round(2)
    df_operations['RECEITA LÍQUIDA'] = df_operations['RECEITA LÍQUIDA'].apply(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    df_final = pd.DataFrame({
        'CLIENTES': df_operations['CLIENTES'],
        'AGENTE': df_operations['AGENTE'],
        'CÓDIGO': df_operations['CÓDIGO_EXTRAIDO'],
        'RECEITA ESTIMADA': df_operations['RECEITA ESTIMADA'],
        'VALOR LÍQUIDO': df_operations['VALOR LÍQUIDO'],
        'DATA DE PAGAMENTO': df_operations['DATA DE PAGAMENTO'],
        'DATA DO RENDIMENTO': df_operations['DATA DO RENDIMENTO'],
        'RECEITA LÍQUIDA': df_operations['RECEITA LÍQUIDA'],
        'MOVIMENTAÇÃO': df_operations['MOVIMENTAÇÃO']
    })

    return df_final

def filtro_agent(df_final, nome_agent):
    if not isinstance(nome_agent, list):
        raise TypeError("nome_agents deve ser uma lista de strings.")
    if not all(isinstance(nome, str) for nome in nome_agent):
        raise ValueError("A lista nome_agents deve conter apenas strings.")
    df_filtrado_agent = df_final[df_final['Assessor'].isin(nome_agent)]
    return df_filtrado_agent[['Agentes', 'Data de Pagamento', 'Data do Rendimento', 'CÓDIGO', 'RECEITA ESTIMADA', 'Valor Líquido', 'Receita Liquida']]

def save_dataframe_to_csv(df, filename, desktop_path):   
    if df is None:
        print(f"Erro: Não foi possível salvar {filename} em CSV, pois o DataFrame é None.")
        return
    filepath = os.path.join(desktop_path, filename)
    try:
        df.to_csv(filepath, index=False, encoding='utf-8-sig', sep=';')
        print(f"DataFrame salvo com sucesso em: {filepath}")
    except Exception as e:
        print(f"Erro ao salvar DataFrame em CSV: {e}")

def main():
    df_final = handling_data()
    df_final_codigo = df_final[df_final['CÓDIGO'].notna()]
    df_receita_final = df_final_codigo[
        df_final_codigo['RECEITA LÍQUIDA'].notna() &
        df_final_codigo['DATA DO RENDIMENTO'].notna()
    ]
    agents_filtro = ["Antonio", "Luiz", "Bruno"] 
    df_final_filtrado = filtro_agent(df_receita_final, agents_filtro)
    if df_final is not None:
        print("\n### Resultado Final ###")
        print(df_final)
        print("Resultado com filtro de códigos válidos:")
        print(df_final_codigo)
        print("Resultado com filtros da Receita Liquida:")
        print(df_receita_final)
        print("Resultado com filtros do(s) agente(s) escolhido(s):")
        print(df_final_filtrado)
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        save_dataframe_to_csv(df_receita_final, "df_receita_final.csv", desktop_path)
        save_dataframe_to_csv(df_final_filtrado, "df_final_filtrado.csv", desktop_path) 

    else:
        print("Erro: O processamento falhou.")

if __name__ == "__main__":
    main()
