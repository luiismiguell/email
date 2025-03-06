import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

def enviar_email(destinatario, assunto, mensagem, remetente, senha):
    try:
        # Configuração do servidor SMTP (Exemplo: Gmail)
        servidor_smtp = "smtp.gmail.com"
        porta_smtp = 587

        # Criando o objeto do e-mail
        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatario
        msg['Subject'] = assunto

        # Adicionando o corpo do e-mail
        msg.attach(MIMEText(mensagem, 'plain'))

        # Conectando ao servidor e enviando o e-mail
        servidor = smtplib.SMTP(servidor_smtp, porta_smtp)
        servidor.starttls()
        servidor.login(remetente, senha)
        servidor.sendmail(remetente, destinatario, msg.as_string())
        servidor.quit()

        print(f"E-mail enviado para {destinatario}")
    except Exception as e:
        print(f"Erro ao enviar e-mail para {destinatario}: {e}")


# Ler a planilha
arquivo_excel = "destinatarios.xlsx"  # Nome do arquivo
planilha = pd.read_excel(arquivo_excel)

# Configuração do remetente
email_remetente = (os.getenv("email_remetente"))  # Coloque seu e-mail
senha_remetente = (os.getenv("senha_remetente"))  # Coloque sua senha (ou utilize autenticação de app)

# Assunto e mensagem padrão
assunto_email = "Assunto do E-mail"
mensagem_email = "Olá, este é um e-mail automático enviado via Python!"

# Enviando e-mails
for index, linha in planilha.iterrows():
    destinatario_email = linha['Email']  # Coluna que contém os e-mails na planilha
    enviar_email(destinatario_email, assunto_email, mensagem_email, email_remetente, senha_remetente)
