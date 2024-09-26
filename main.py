# Import Cliente
import Cliente

# Caminhos dos arquivos
input_filepath = 'C:/Users/Vinishow/Downloads/Formulario_Respostas.xlsx'
output_filepath = 'C:/Users/Vinishow/Downloads/Clientes.xlsx'

clientes = Cliente.read_from_excel(input_filepath)

if clientes:
    Cliente.write_to_excel(clientes, output_filepath)
    print("Arquivo criado com sucesso!")
else:
    print("Nenhum cliente encontrado.")
