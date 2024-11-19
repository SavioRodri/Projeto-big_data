import pandas as pd
import os

i = 2024

while i >= 2017:
    diretorio = 'Tabelas/Mortalidade Leishmaniose'
    arquivos = os.listdir(diretorio)
    print(arquivos)
    print(f'-------------------inicio---------------------{i}\n')
    for arqsep in arquivos:
        print(arqsep)
        caminho_arquivo = f'Tabelas/Mortalidade Leishmaniose/{arqsep}'
        excel_file = pd.ExcelFile(caminho_arquivo)

        for aba in excel_file.sheet_names:
            df = pd.read_excel(caminho_arquivo, sheet_name=aba)
            df['Ano'] = i
            df['Nome_Aba'] = aba
            print(df.head())
    print(f'-------------------final---------------------{i}\n')
    i -= 1