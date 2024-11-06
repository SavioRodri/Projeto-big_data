import pandas as pd
import os

i = 2024

while i >= 2017:
    diretorio = (f'Dados Leichmaniose/Mortalidade/BA/{i}')
    arquivos = os.listdir(diretorio)
    print(arquivos)
    print(f'-------------------inicio---------------------{i}\n')
    for arqsep in arquivos:
        print(arqsep)
        diretorio = (f'Dados Leichmaniose/Mortalidade/BA/{i}/{arqsep}')
        df = pd.read_csv(diretorio, sep=';', encoding='ISO-8859-1', header=0, skiprows=3, skipfooter=14, engine='python')
        print(df.head())
    print(f'-------------------final---------------------{i}\n')
    i = i - 1
