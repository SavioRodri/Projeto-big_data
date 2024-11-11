import pandas as pd
import os

class LeitorCSV:
    def __init__(self, diretorio_base, ano_final, ano_inicial):
        self.diretorio_base = diretorio_base
        self.ano_final = ano_final
        self.ano_inicial = ano_inicial

    def _processar_arquivo(self, ano, arquivo):
        diretorio_arquivo = os.path.join(self.diretorio_base, str(ano), arquivo) # Construindo diretório
        
        # Tentando ler o arquivo CSV
        try:
            dataframe = pd.read_csv(diretorio_arquivo,
                                    sep=';',
                                    encoding='ISO-8859-1',
                                    header=0,
                                    skiprows=3,
                                    skipfooter=14,
                                    engine='python')
            print(f"Arquivo: {arquivo}")
            print(dataframe.head())
        except Exception as e:
            print(f"Erro ao processar o arquivo {arquivo}: {e}")

    def processar_arquivos(self):
        # Itera pelas pastas de ano
        for ano in range(self.ano_inicial, self.ano_final+1, +1):

            diretorio_ano = os.path.join(self.diretorio_base, str(ano))
            if not os.path.exists(diretorio_ano):
                print(f"Diretório {diretorio_ano} não encontrado.")
            
            arquivos = os.listdir(diretorio_ano)
            print(f'-------------------início---------------------{ano}\n')
            
            # Processa cada arquivo dentro do ano
            for arquivo in arquivos:
                self._processar_arquivo(ano, arquivo)
            
            print(f'-------------------final---------------------{ano}\n')
