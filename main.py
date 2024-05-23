import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# Função para enviar o e-mail
def send_email(from_email, to_email, subject, body, attachment_filename):
    smtp_server = 'mail.example.com'
    smtp_port = 540
    smtp_username = 'email@example.com'
    smtp_password = 'tua_senha'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {attachment_filename}')
    msg.attach(part)

    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(smtp_username, smtp_password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

# Função para ajustar automaticamente a largura das colunas
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

# Função para converter CSV para Excel e ajustar as colunas
def csv_to_excel(csv_filename, excel_filename):
    df = pd.read_csv(csv_filename)
    book = Workbook()
    sheet = book.active

    for r in dataframe_to_rows(df, index=False, header=True):
        sheet.append(r)

    book = auto_adjust_column_width(book)
    book.save(excel_filename)

# Função para adicionar impressoras e toners
def add_impressora_toner():
    print("1. Adicionar Impressora")
    print("2. Adicionar Toner")
    escolha = input("Escolha uma opção: ")
    
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
    else:
        print("Opção inválida!")

# Função para atualizar a tabela
def update_table():
    try:
        df = pd.read_csv('relatorio_toner.csv')
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Toner', 'Quantidade', 'Destino', 'Impressora'])
    
    for _, row in df.iterrows():
        print(row.to_string())
    
    toner = input("Digite o nome do toner a ser atualizado: ")
    quantidade = input("Digite a nova quantidade: ")
    destino = input("Digite o novo destino: ")
    impressora = input("Digite a nova impressora: ")

    df.loc[df['Toner'] == toner, ['Quantidade', 'Destino', 'Impressora']] = [quantidade, destino, impressora]
    df.to_csv('relatorio_toner.csv', index=False)
    print("Tabela atualizada com sucesso!")

# Função para quantificar toners
def quantificar_toners():
    try:
        df = pd.read_csv('relatorio_toner.csv')
    except FileNotFoundError:
        print("Nenhum relatório encontrado.")
        return

    df['Quantidade'] = pd.to_numeric(df['Quantidade'], errors='coerce').fillna(0).astype(int)
    print(df.groupby('Toner')['Quantidade'].sum())

# Função para relacionar toners a impressoras
def relacionar_toners_impressoras():
    try:
        df = pd.read_csv('relatorio_toner.csv')
    except FileNotFoundError:
        print("Nenhum relatório encontrado.")
        return
    
    impressora = input("Digite o nome da impressora: ")
    related_toners = df[df['Impressora'] == impressora]['Toner'].unique()
    print(f"Toners relacionados à impressora {impressora}: {', '.join(related_toners)}")

# Função para enviar a tabela por e-mail
def enviar_tabela_email():
    from_email = 'matheus.torres@reals.com.br'
    to_email = input("Digite o e-mail do destinatário: ")
    subject = 'Relatório de Toner'
    body = 'Por favor, encontre em anexo o relatório de toner.'
    relatorio_filename = 'relatorio_toner.xlsx'

    # Converte o CSV para Excel antes de enviar
    csv_to_excel('relatorio_toner.csv', relatorio_filename)

    send_email(from_email, to_email, subject, body, relatorio_filename)

# Menu principal
def main_menu():
    while True:
        print("Menu Principal")
        print("1. Adicionar Impressora e Toner")
        print("2. Atualizar tabela")
        print("3. Quantificar toners")
        print("4. Relacionar Toners a Impressoras")
        print("5. Enviar tabela para email")
        print("6. Sair")
        
        escolha = input("Escolha uma opção: ")
        
        if escolha == "1":
            add_impressora_toner()
        elif escolha == "2":
            update_table()
        elif escolha == "3":
            quantificar_toners()
        elif escolha == "4":
            relacionar_toners_impressoras()
        elif escolha == "5":
            enviar_tabela_email()
        elif escolha == "6":
            print("Saindo...")
            break
        else:
            print("Opção inválida! Tente novamente.")

if __name__ == "__main__":
    main_menu()
