import pandas as pd

# colocar o path final da variavel aqui
file_path = ""

data_frame = pd.read_excel(file_path)
data_frame.to_csv("respostas.csv", index=False)