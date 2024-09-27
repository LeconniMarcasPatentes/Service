import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
from email.mime.base import MIMEBase
from email import encoders

# Carregar variáveis de ambiente do arquivo .env
load_dotenv('email.env')

def enviar_email(destinatario, assunto, corpo, arquivo_anexo_path=None):
  try:
    remetente = os.getenv('EMAIL_HOSTNET')  
    senha = os.getenv('SENHA_HOSTNET') 

    # Configurar servidor SMTP
    servidor = smtplib.SMTP('smtp.hostnet.com.br', 587)
    servidor.starttls()
    servidor.login(remetente, senha)

    # Criar o corpo do email
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    # Anexar o corpo da mensagem
    msg.attach(MIMEText(corpo, 'plain'))

    # Anexar arquivo
    if arquivo_anexo_path:
      with open(arquivo_anexo_path, 'rb') as anexo:
        parte_arq = MIMEBase('application', 'octet-stream')
        parte_arq.set_payload(anexo.read())
        encoders.encode_base64(parte_arq)
        parte_arq.add_header('Content-Disposition', f'attachment; filename={os.path.basename(arquivo_anexo_path)}')
        msg.attach(parte_arq)

    # Enviar o email
    servidor.send_message(msg)
    servidor.quit()
    print("E-mail enviado com sucesso!")
  except Exception as e:
    print(f"Erro ao enviar o e-mail: {e}")

# Exemplo de uso
#email do vendedor
destinatario = "giuseppe.cordeiro@gmail.com"

assunto = "Email referente ao contrato de MARCA"
corpo = "Segue em anexo o contrato preenchido com as informações do cliente."
# arquivo_anexo_path = "teste.txt"

enviar_email(destinatario, assunto, corpo)
