from planilha_module import Planilha
from cliente_module import Cliente
from vendedor_module import Vendedor
from utils import *

forms_filepath = "data/raw/Respostas_Fomulario.xlsx"

try:
    
    respostas = Planilha.ler_formulario( forms_filepath )
    
    Cliente.gravar_planilha_clientes( respostas )
    Vendedor.gravar_planilha_vendedores( respostas )

    # Cliente.gerar_contratos( clientes )

except Exception as e:
    print( f"\nOcorreu um erro ao processar o arquivo: {e}\n" )