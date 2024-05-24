from file_utils import add_impressora_toner, update_table, quantificar_toners, relacionar_toners_impressoras, visualizar_relatorio, csv_to_excel
from email_utils import send_email
from datetime import datetime

def enviar_tabela_email():
    from_email = 'matheus.torres@reals.com.br'
    to_email = input("Digite o e-mail do destinatário: ")
    subject = 'Relatório de Toner'
    body = 'Por favor, encontre em anexo o relatório de toner.'
    data_atual = datetime.now().strftime('%Y-%m-%d')
    relatorio_filename = f'relatorio_toner_{data_atual}.xlsx'

    # Converte o CSV para Excel antes de enviar
    csv_to_excel('relatorio_toner.csv', relatorio_filename)

    send_email(from_email, to_email, subject, body, relatorio_filename)

def main_menu():
    while True:
        print("================================")
        print("Menu Principal")
        print("1. Adicionar Impressora e Toner")
        print("2. Atualizar tabela")
        print("3. Quantificar toners")
        print("4. Relacionar Toners a Impressoras")
        print("5. Enviar tabela para email")
        print("6. Visualizar Relatório")
        print("7. Sair")
        print("================================")

        while True:
            escolha = input("Escolha uma opção: ")
            if escolha in ["1", "2", "3", "4", "5", "6", "7"]:
                break
            else:
                print("Opção inválida! Tente novamente.")

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
            visualizar_relatorio()
        elif escolha == "7":
            print("Saindo...")
            break
