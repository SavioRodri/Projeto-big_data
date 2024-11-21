from leitor_csv import LeitorCSV

if __name__ == "__main__":
    diretorio_base = r'C:\Users\lipe_\OneDrive - Educacional\projeto_big_data-estacio\Dados Leichmaniose\Mortalidade\BA'
    processor = LeitorCSV(diretorio_base, ano_inicial=2017, ano_final=2024)
    processor.processar_arquivos()

