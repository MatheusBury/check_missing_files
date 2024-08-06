import os
import pandas as pd

def get_files_in_directory(directory):
    """
    Obtém a lista de arquivos em um diretório.
    """
    return set(os.listdir(directory))

def get_files_from_csv(excel_file):
    """
    Lê os nomes dos arquivos a partir de um arquivo CSV.
    """
    df = pd.read_excel(excel_file)
    return set(df['Picture'])

def find_missing_files(directory, excel_file, output_excel):
    """
    Compara arquivos em um diretório com os arquivos listados em um CSV e 
    escreve os arquivos ausentes em um arquivo Excel.
    """
    files_in_directory = get_files_in_directory(directory)
    files_from_csv = get_files_from_csv(excel_file)
    
    missing_files = files_from_csv - files_in_directory
    
    if missing_files:
        df_missing = pd.DataFrame(list(missing_files), columns=['Missing Files'])
        df_missing.to_excel(output_excel, index=False)
        print(f"Arquivo Excel '{output_excel}' criado com os arquivos ausentes.")
    else:
        print("Não há arquivos ausentes.")

if __name__ == "__main__":
    directory_path = r'G:\Drives compartilhados\Modec\MV33\imagens OS'
    excel_file_path = r"C:\Users\Matheus\Downloads\mv33_image.xlsx"
    output_excel_path = r'F:\saida.xlsx'
    
    find_missing_files(directory_path, excel_file_path, output_excel_path)
