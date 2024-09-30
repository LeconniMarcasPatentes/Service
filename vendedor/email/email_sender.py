import win32com.client as win32
import vendedor.email

# Inicializa o Outlook
outlook = win32.Dispatch("Outlook.Application")

# Lista de destinatários (com domínio leconni.com)
emails = "marconnirodrigues@leconni.com"

# Cria o e-mail
message = outlook.CreateItem(0)
message.Display()  # Exibe o e-mail antes de enviar (pode ser removido se não quiser exibir)

# Configuração de destinatários
message.To = vendedor.email  # Para
message.Bcc = emails  # Cópia oculta

# Assunto e corpo do e-mail
message.Subject = "Mensagem"
message.Body = "Testes de envio"

# Corpo em HTML com uma imagem
body = """
OIIII SHOW
"""

# Define o corpo como HTML
message.HTMLBody = body

# Envia o e-mail
message.Send()
print("email enviado")