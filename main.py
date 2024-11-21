from leitor_csv import LeitorCSV

if __name__ == "__main__":
    diretorio_base = r'D:\Users\Lipe_\Área de Trabalho\SOS'
    pastas = [str(i) for i in range(1, 9)]  # Pastas de 1 a 8
    processor = LeitorCSV(diretorio_base, pastas)
    processor.processar_arquivos()
