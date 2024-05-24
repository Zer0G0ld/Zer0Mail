import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime

def auto_adjust_column_width(book):
    sheet = book.active
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column].width = adjusted_width
    return book

def csv_to_excel(csv_filename, excel_filename):
    df = pd.read_csv(csv_filename)
    book = Workbook()
    sheet = book.active

    for r in dataframe_to_rows(df, index=False, header=True):
        sheet.append(r)

    book = auto_adjust_column_width(book)
    book.save(excel_filename)

def add_impressora_toner():
    print("1. Adicionar Impressora")
    print("2. Adicionar Toner")
    while True:
        escolha = input("Escolha uma opção: ")
        if escolha in ["1", "2"]:
            break
        else:
            print("Opção inválida! Tente novamente.")

    if escolha == "1":
        impressora = input("Digite o nome da impressora: ")
        with open("impressoras.txt", "a") as file:
            file.write(impressora + "\n")
        print(f"Impressora {impressora} adicionada com sucesso!")
    elif escolha == "2":
        toner = input("Digite o nome do toner: ")
        with open("toners.txt", "a") as file:
            file.write(toner + "\n")
        print(f"Toner {toner} adicionado com sucesso!")

def update_table(toner, quantidade, destino, impressora):
    try:
        df = pd.read_csv('relatorio_toner.csv')
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Toner', 'Quantidade', 'Destino', 'Impressora', 'Data'])

    # Mostrar o relatório atual
    print(df.to_string(index=False))

    data_atual = datetime.now().strftime('%Y-%m-%d')
    df.loc[df['Toner'] == toner, ['Quantidade', 'Destino', 'Impressora', 'Data']] = [quantidade, destino, impressora, data_atual]
    df.to_csv('relatorio_toner.csv', index=False)
    print("Tabela atualizada com sucesso!")

def quantificar_toners():
    try:
        df = pd.read_csv('relatorio_toner.csv')
    except FileNotFoundError:
        print("Nenhum relatório encontrado.")
        return None

    df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce').fillna(0).astype(int)
    return df.groupby('Toner')['Quantidade'].sum()

def relacionar_toners_impressoras(impressora):
    try:
        df = pd.read_csv('relatorio_toner.csv')
    except FileNotFoundError:
        print("Nenhum relatório encontrado.")
        return None
    
    related_toners = df[df['Impressora'] == impressora]['Toner'].unique()
    return related_toners

def visualizar_relatorio():
    try:
        df = pd.read_csv('relatorio_toner.csv')
        return df
    except FileNotFoundError:
        print("Nenhum relatório encontrado.")
        return None
