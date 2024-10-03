from google.oauth2 import service_account
from googleapiclient.discovery import build

# Carregar as credenciais
SCOPES = ['https://www.googleapis.com/auth/drive.metadata.readonly']
SERVICE_ACCOUNT_FILE = 'converter/a.json'

credentials = service_account.Credentials.from_service_account_file(
  SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Criar o serviço do Google Drive
service = build('drive', 'v3', credentials=credentials)

# Função para listar arquivos e obter o caminho
def get_file_path(file_id):
  # Obter metadados do arquivo
  file = service.files().get(fileId=file_id, fields='name, parents').execute()
  file_name = file.get('name')
  parents = file.get('parents')
  
  # Construir o caminho
  path = file_name
  if parents:
    # Obter o nome da pasta pai
    for parent_id in parents:
      parent = service.files().get(fileId=parent_id, fields='name, parents').execute()
      path = parent.get('name') + '/' + path
      # Repetir para pais até a raiz
      while parent.get('parents'):
        parent_id = parent.get('parents')[0]
        parent = service.files().get(fileId=parent_id, fields='name, parents').execute()
        path = parent.get('name') + '/' + path

  return path

# Exemplo de uso
file_id = '1aCFZ0Xudn72J898tOHijY5skdBvoZZmJYEa1yRhJL4A'
file_path = get_file_path(file_id)
print(f'File path: {file_path}')
