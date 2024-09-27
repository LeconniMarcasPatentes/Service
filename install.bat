@echo off
REM Instalar dependências
pip install pandas
pip install pydrive
pip install openpyxl
pip install secure-smtplib
pip install python-dotenv

REM Executar o programa
cls
python main.py
