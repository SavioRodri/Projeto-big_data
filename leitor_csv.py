import pandas as pd
import os

class LeitorCSV:
    def __init__(self, diretorio_base, ano_final, ano_inicial):
        self.diretorio_base = diretorio_base
        self.ano_final = ano_final
        self.ano_inicial = ano_inicial

    def _converter_xlsx(self, arquivo, dataframe, writer):
        """ Cria uma aba no arquivo Excel para cada DataFrame
            usando o nome do arquivo (ou parte dele) para nomear a aba.

        Args:
            arquivo (_type_): _description_
            dataframe (_type_): _description_
            writer (_type_): _description_
        """
        nome_aba = arquivo[:31]  # Limitar o nome da aba para 31 caracteres
        dataframe.to_excel(writer, sheet_name=nome_aba, index=True)

    def _processar_arquivo(self, ano, arquivo, writer):
        """ Cria um DataFrame baseado no arquivo.csv.

        Args:
            ano (_type_): _description_
            arquivo (_type_): _description_
            writer (_type_): _description_
        """
        diretorio_arquivo = os.path.join(self.diretorio_base, str(ano), arquivo) # Construindo o caminho do arquivo
        
        # Tentando ler o arquivo CSV
        try:
            dataframe = pd.read_csv(diretorio_arquivo,
                                    sep=';',
                                    encoding='ISO-8859-1',
                                    header=0,
                                    skiprows=3,
                                    skipfooter=14,
                                    engine='python')
            self._converter_xlsx(arquivo, dataframe, writer)
        except Exception as e:
            print(f"Erro ao processar o arquivo {arquivo}: {e}")

    def processar_arquivos(self):
        """ Percorre a pasta fornecida em busca de arquivos com as
            respectivas datas.
        """
        # Criando o diretório Excel, caso não exista
        diretorio_xlsx = 'Excel'
        if not os.path.exists(diretorio_xlsx):
            os.makedirs(diretorio_xlsx)
        
        # Itera pelas pastas de ano
        for ano in range(self.ano_inicial, self.ano_final + 1):
            diretorio_ano = os.path.join(self.diretorio_base, str(ano))
            
            if not os.path.exists(diretorio_ano):
                print(f"Diretório {diretorio_ano} não encontrado.")
                continue  # Pula para o próximo ano se o diretório não existir
            
            # Cria o ExcelWriter uma vez para cada ano
            caminho_arquivo = os.path.join(diretorio_xlsx, f'período_{ano}.xlsx')
            with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                
                # Lista todos os arquivos no diretório do ano
                arquivos = os.listdir(diretorio_ano)
                
                # Processa cada arquivo dentro do ano
                for arquivo in arquivos:
                    self._processar_arquivo(ano, arquivo, writer)

            print(f'Arquivo Excel para o ano {ano} gerado com sucesso em {caminho_arquivo}.')
