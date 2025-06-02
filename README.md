# Limpeza, leitura e organização de dados (ETL)

## Sumário

- [Descrição da Aplicação](#descrição-da-aplicação)
- [Objetivo](#objetivo)
- [Como Usar](#como-usar)
  - [Pré-requisitos](#pré-requisitos)
  - [Instalação](#instalação)
- [Contribuições](#contribuições)

## Descrição da Aplicação

Este projeto é um simulador desenvolvido em Python que modela e analisa as trajetórias de veículos movidos a energias renováveis, como solar e eólica. A aplicação permite aos usuários avaliar o desempenho desses veículos sob diferentes condições ambientais e de operação, fornecendo insights sobre eficiência energética e viabilidade de rotas.

## Objetivo

O objetivo principal desta aplicação é fornecer uma ferramenta que limpa e organiza arquivos em CSV, XLSM e XLSX, onde:

- Extrai os arquivos verificando se possuem algum erro de compactabilidade ou se estão corrompidos;
- Cria DataFrames com os dados e realiza limpezas básicas;
- Realiza junções no caso especifico de 'bases_revenue', onde se encontra mais de uma tabela de dados na mesma planilha, transformando em um unico dataframe;
- Tranforma e elebara a pesquisa das colunas para serem utilizadas no programa principal;
- Personaliza todas as planilhas utilizando merges entre elas e registrando as datas de acordo com o nome especifico de cada planilha que possuir esta informação;
- Por fim, tendo a tabela personalizada da união entre 3 tipos de arquivos com numeros diferentes de informações diarias.

## Como Usar

### Pré-requisitos

- **Python 3.x instalado:** Verifique se o Python está instalado executando `python --version` ou `python3 --version` no terminal. Caso não esteja instalado, faça o download e instale a versão mais recente do Python a partir do [site oficial](https://www.python.org/downloads/).

- **pip instalado:** O `pip` geralmente é instalado junto com o Python. Para verificar, execute `pip --version` ou `pip3 --version`. Se não estiver instalado, siga as instruções na [documentação oficial](https://pip.pypa.io/en/stable/installation/) para instalá-lo.

### Instalação

1. **Clone o repositório:**

   ```bash
   git clone https://github.com/seu-usuario/simulador-veiculos-renovaveis.git
   cd 

2. **Instale as dependêcias:**

   ```bash
   pip install -r requirements.txt

3. **Faça um primeiro teste e veja as colunas e informações necessárias que são especificas dos dados que for utilizado no programa principal.**

4. **Por fim, realize as alterações necessárias utilizando como base as bibliotecas da API**

### Contribuições

- **Este código foi generalizado para casos que necessitem de mais de um arquivo para verificações de receita entre diferentes bases e produtos de uma empresa.**

### Observação

Muitas vezes alguns arquivos possuem senhas para acesso, principalmente quando são dados comprados de terceiros. Logo pode estar implementando na leitura dos dados a seguinte função:

```bash
import msoffcrypto
import pandas as pd
from io import BytesIO

    def ler_excel_protegido(caminho_arquivo, senha):
        with open(caminho_arquivo, 'rb') as arquivo:
            arquivo_descriptografado = msoffcrypto.OfficeFile(arquivo)
            arquivo_descriptografado.load_key(password=senha)
            arquivo_decifrado = BytesIO()
            arquivo_descriptografado.decrypt(arquivo_decifrado)
            df = pd.read_excel(arquivo_decifrado, engine='openpyxl')
        return df

