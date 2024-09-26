import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

# Autenticar e acessar a planilha
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file('credenciais/key_gdrive_api.json', scopes=scope)
client = gspread.authorize(creds)

# Abrir a planilha pelo nome ou URL
sheet = client.open('https://docs.google.com/spreadsheets/d/1aCFZ0Xudn72J898tOHijY5skdBvoZZmJYEa1yRhJL4A/edit?usp=drive_web&ouid=117929815860061714751').sheet1  # Se tiver várias abas, use sheet1, sheet2, etc.

# Ler os dados
data = sheet.get_all_records()

# Converter para DataFrame do pandas
df = pd.DataFrame(data)

print(df.head())

# ERRO DE FORMATAÇÃO DO JSON