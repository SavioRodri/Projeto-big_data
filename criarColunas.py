import pandas as pd
import os
import re

# Definir o diretório onde os arquivos finalizados serão salvos
diretorio_saida = r'D:\Users\Lipe_\Área de Trabalho\SOS\Dados finalizados'
if not os.path.exists(diretorio_saida):
    os.makedirs(diretorio_saida)

# Definir o diretório de entrada
diretorio = r'D:/Users/Lipe_/Área de Trabalho/SOS/Dados processados'
arquivos = os.listdir(diretorio)
print(arquivos)

print(f'-------------------inicio---------------------\n')

# Processar cada arquivo
for arqsep in arquivos:
    print(arqsep)
    caminho_arquivo = os.path.join(diretorio, arqsep)  # Usar os.path.join para construir o caminho do arquivo
    excel_file = pd.ExcelFile(caminho_arquivo)

    # Extrair o ano do nome do arquivo usando regex
    match = re.match(r'(\d{4})', arqsep)
    if match:
        ano = int(match.group(1))
    else:
        ano = "Desconhecido"
        print(f"Não foi possível extrair o ano do arquivo: {arqsep}")
    
    # Definir o nome e caminho do novo arquivo Excel
    nome_novo_arquivo = f"{os.path.splitext(arqsep)[0]}_{ano}.xlsx"
    caminho_arquivo_saida = os.path.join(diretorio_saida, nome_novo_arquivo)
    
    # Salvar todas as abas em um único arquivo Excel
    with pd.ExcelWriter(caminho_arquivo_saida, engine='openpyxl') as writer:
        for aba in excel_file.sheet_names:
            df = pd.read_excel(caminho_arquivo, sheet_name=aba)
            
            # Adicionar colunas Ano e Nome_Aba
            df['Ano'] = ano
            df['Nome_Aba'] = aba
            print(df.head())
            
            # Salvar o DataFrame em uma nova aba do arquivo Excel
            df.to_excel(writer, index=False, sheet_name=aba)

print(f'-------------------final---------------------\n')
