import pandas as pd

def mesclar_abas(arquivo_entrada, aba_saida):
    # Lê todas as abas da planilha de entrada
    xls = pd.ExcelFile(arquivo_entrada)
    todas_abas = []

    # Percorre todas as abas e adiciona ao dataframe
    for aba in xls.sheet_names:
        df = pd.read_excel(arquivo_entrada, sheet_name=aba)
        todas_abas.append(df)

    # Concatena todos os dataframes em um único dataframe
    resultado = pd.concat(todas_abas, ignore_index=True)

    # Salva o dataframe resultante em uma nova aba
    with pd.ExcelWriter(arquivo_entrada, engine='openpyxl', mode='a') as writer:
        resultado.to_excel(writer, sheet_name=aba_saida, index=False)

    print(f'Nova aba "{aba_saida}" criada com sucesso em {arquivo_entrada}.')

# Exemplos de uso
arquivo_entrada = r'D:\Users\Lipe_\Área de Trabalho\SOS\Dados finalizados\MorteRes Sexo segundo Mun-BA_Desconhecido.xlsx'
aba_saida = 'Planilha_Mesclada'
mesclar_abas(arquivo_entrada, aba_saida)
