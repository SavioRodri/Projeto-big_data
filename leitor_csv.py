import pandas as pd
import os
import re

class LeitorCSV:
    def __init__(self, diretorio_base, pastas):
        self.diretorio_base = diretorio_base
        self.pastas = pastas

    def _converter_xlsx(self, arquivo, dataframe, writer):
        """ Cria uma aba no arquivo Excel para cada DataFrame
            usando o nome do arquivo (ou parte dele) para nomear a aba.

        Args:
            arquivo (str): Nome do arquivo CSV.
            dataframe (pd.DataFrame): DataFrame contendo os dados do arquivo CSV.
            writer (pd.ExcelWriter): Objeto ExcelWriter para escrever os dados no Excel.
        """
        nome_arquivo = os.path.splitext(arquivo)[0]  # Remove extensão
        nome_aba = nome_arquivo[:31]  # Limitar o nome da aba para 31 caracteres
        dataframe.to_excel(writer, sheet_name=nome_aba, index=True)

    def _processar_arquivo(self, arquivo, writer, pasta):
        """ Cria um DataFrame baseado no arquivo CSV e extrai o ano do nome do arquivo.

        Args:
            arquivo (str): Nome do arquivo CSV.
            writer (pd.ExcelWriter): Objeto ExcelWriter para escrever os dados no Excel.
            pasta (str): Nome da pasta sendo processada.
        """
        diretorio_arquivo = os.path.join(self.diretorio_base, pasta, arquivo)  # Construindo o caminho do arquivo
        
        # Tentando extrair o ano do nome do arquivo
        match = re.match(r'(\d{4})', arquivo)
        if match:
            ano = int(match.group(1))
        else:
            print(f"Não foi possível extrair o ano do arquivo {arquivo}")
            return

        # Tentando ler o arquivo CSV
        try:
            dataframe = pd.read_csv(diretorio_arquivo,
                                    sep=';',
                                    encoding='ISO-8859-1',
                                    header=0,
                                    skiprows=3,
                                    skipfooter=14,
                                    engine='python')
            dataframe['Ano'] = ano
            dataframe['Pasta'] = pasta
            self._converter_xlsx(arquivo, dataframe, writer)
        except Exception as e:
            print(f"Erro ao processar o arquivo {arquivo} na pasta {pasta}: {e}")

    def processar_arquivos(self):
        """ Percorre as pastas fornecidas em busca de arquivos CSV e os converte para arquivos Excel. """
        # Criando o diretório Excel, caso não exista
        diretorio_xlsx = 'Excel'
        if not os.path.exists(diretorio_xlsx):
            os.makedirs(diretorio_xlsx)
        
        # Itera pelas pastas principais (1 a 8)
        for pasta in self.pastas:
            print(f'Processando arquivos na pasta {pasta}')
            caminho_pasta_saida = os.path.join(diretorio_xlsx, f'pasta_{pasta}.xlsx')
            arquivos_existem = False
            
            # Lista todos os arquivos no diretório da pasta
            diretorio_pasta = os.path.join(self.diretorio_base, pasta)
            arquivos = os.listdir(diretorio_pasta)
            
            if arquivos:
                arquivos_existem = True

            # Criar o ExcelWriter somente se houver arquivos a processar
            if arquivos_existem:
                with pd.ExcelWriter(caminho_pasta_saida, engine='openpyxl') as writer:
                    for arquivo in arquivos:
                        self._processar_arquivo(arquivo, writer, pasta)

                print(f'Arquivo Excel para a pasta {pasta} gerado com sucesso em {caminho_pasta_saida}.')
