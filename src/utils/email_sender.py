import os
import win32com.client as win32

def gerar_corpo_email(vendedor_nome, tipo_contrato, nome_empresa, servico_prestado):
  """
  Gera o corpo do e-mail em formato HTML com as informações do vendedor e do contrato.

  Parameters:
    vendedor_nome (str): Nome do vendedor.
    tipo_contrato (str): Tipo do contrato.
    nome_empresa (str): Nome da empresa.
    servico_prestado (str): Serviço prestado.
    imagem_assinatura (str): Caminho da imagem da assinatura.

  Returns:
    str: Corpo do e-mail em formato HTML.
  """
  
  corpo_html = f"""
    <html>
      <body>
        <p>Olá {vendedor_nome},</p>
        <p>Segue em anexo o contrato <strong>{tipo_contrato}</strong> referente à empresa <strong>{nome_empresa}</strong> sobre <strong>{servico_prestado}</strong>.</p>
        <p>Por favor, revise o documento e, se tudo estiver correto, encaminhe-o para o cliente.</p>
        <p>Atenciosamente,</p>
        <p><strong>Nome da Sua Empresa</strong></p>
        <img src="cid:assinatura" alt="Assinatura" style="width: 200px; height: auto;" />
      </body>
    </html>
    """
  return corpo_html

def pegar_contrato(nome_empresa):
  """
  Retorna o caminho do contrato referente à empresa, assumindo que o nome do arquivo é baseado no nome da empresa.
  Parameters:
    nome_empresa (str): Nome da empresa.
  Returns:
  str: Caminho completo para o arquivo de contrato.
  """
  
  
  # Substituir espaços por underscores no nome da empresa
  nome_arquivo = f"contrato_{nome_empresa.replace(' ', '_')}.docx"
  
  # Caminho completo para o arquivo de contrato (ajuste o diretório conforme necessário)
  caminho_arquivo = os.path.join('data', 'contatos', nome_arquivo)

  return caminho_arquivo

def enviar_email_contrato(vendedor_email, vendedor_nome, tipo_contrato, nome_empresa, servico_prestado):
  """
  Envia um e-mail com o contrato de venda para o vendedor e CCO para o domínio específico.

  Parameters:
    vendedor_email (str): O e-mail do vendedor.
    vendedor_nome (str): Nome do vendedor.
    tipo_contrato (str): Tipo do contrato.
    nome_empresa (str): Nome da empresa.
    servico_prestado (str): Serviço prestado.
  """
  # Inicializa o Outlook
  outlook = win32.Dispatch("Outlook.Application")

  # Lista de destinatários ocultos
  emails = "marconnirodrigues@leconni.com"

  # Cria o e-mail
  message = outlook.CreateItem(0)

  # Configuração de destinatários
  message.To = vendedor_email  # Para
  message.Bcc = emails  # Cópia oculta

  # Gera o corpo do e-mail
  corpo_html = gerar_corpo_email(vendedor_nome, tipo_contrato, nome_empresa, servico_prestado)

  # Assunto do e-mail
  message.Subject = f"Contrato de venda: {nome_empresa}"

  # Define o corpo como HTML
  message.HTMLBody = corpo_html

  # Caminho da imagem da assinatura
  imagem_assinatura = os.path.join('img', 'assinatura_contato.png')

  # Caminho para o contrato
  contrato_path = pegar_contrato(nome_empresa)

  # Adicionar o contrato como anexo
  message.Attachments.Add(contrato_path)

  # Adiciona a imagem como anexo e referencia na assinatura
  message.Attachments.Add(imagem_assinatura, 1, 0, "assinatura") 

  # Envia o e-mail
  message.Send()
  print("E-mail enviado")