import pandas as pd
import os
import xlsxwriter


directory = 'Projeto-big_data\Tabelas\Mortalidade Leishmaniose'

def combine_sheets_to_excel(directory, output_file):
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):
            try:
                df = pd.read_csv(os.path.join(directory, filename), encoding='utf-8', on_bad_lines='skip')
            except UnicodeDecodeError:
                df = pd.read_csv(os.path.join(directory, filename), encoding='latin1', on_bad_lines='skip')
            sheet_name = os.path.splitext(filename)[0][:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        elif filename.endswith('.xlsx'):
            xls = pd.ExcelFile(os.path.join(directory, filename))
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                truncated_sheet_name = f"{os.path.splitext(filename)[0]}_{sheet_name}"[:31]
                df.to_excel(writer, sheet_name=truncated_sheet_name, index=False)


    writer.close()


"""directory = 'Dados Leichmaniose/Mortalidade/BA/2017'"""
output_file = 'planilhas_combinadas_2017.xlsx' 

combine_sheets_to_excel(directory, output_file)