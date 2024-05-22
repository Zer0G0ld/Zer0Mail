import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Função para enviar o e-mail
def send_email(from_email, to_email, subject, body, attachment_filename):
    # Configuração do servidor SMTP
    smtp_server = 'mail.example.com'
    smtp_port = 540
    smtp_username = 'email@examople.com'
    smtp_password = 'tua_senha'

    # Criando objeto do tipo MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Adicionando o corpo do e-mail
    msg.attach(MIMEText(body, 'plain'))

    # Anexando o arquivo
    with open(attachment_filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {attachment_filename}')
    msg.attach(part)

    # Enviando o e-mail
    try:
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(smtp_username, smtp_password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        print("E-mail enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")

# Lista de modelos de toner
modelos_toner = ['TN1060', 'W1105A C/CHIP', 'TN B021']

# Perguntando ao usuário a quantidade de toner para cada modelo
quantidades = []
for modelo in modelos_toner:
    quantidade = input(f"Quantidade de toner para o modelo {modelo}: ")
    quantidades.append(quantidade)

# Criando um DataFrame pandas com os dados
dados = {'Toner': modelos_toner, 'Quantidade': quantidades}
df = pd.DataFrame(dados)

# Salvando o DataFrame como CSV
relatorio_filename = 'relatorio_toner.csv'
df.to_csv(relatorio_filename, index=False)
print("DataFrame salvo como CSV")

# Enviando o e-mail com o relatório anexado
from_email = 'teu_email@examople.com'
to_email = 'email@examople.com'
subject = 'Relatório de Toner'
body = 'Por favor, encontre em anexo o relatório de toner.'

send_email(from_email, to_email, subject, body, relatorio_filename)
