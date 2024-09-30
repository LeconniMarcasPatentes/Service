# Import Cliente
import Cliente

# Caminhos dos arquivos
input_filepath = 'C:/Users/Vinishow/Downloads/Formulario_Respostas.xlsx'
output_filepath = 'C:/Users/Vinishow/Downloads/Clientes.xlsx'
output_dir = 'C:/Users/Vinishow/Downloads/Contrato-Teste.docx'
template_path = 'C:/Users/Vinishow/Downloads/MODELO_CONTRATO.docx'

clientes = Cliente.read_from_excel(input_filepath)

# if clientes:
#     Cliente.write_to_excel( clientes, output_filepath )
#     print("Planilha criada com sucesso!")
# else:
#     print("Nenhum cliente encontrado.")
    
cliente = clientes[0]
if cliente:
    Cliente.write_to_docs( template_path, output_dir, cliente )
    print("Contrato criado com sucesso.")
else:
    print("Cliente não existe")
