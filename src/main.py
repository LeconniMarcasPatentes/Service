from planilha_module import Planilha

forms_filepath = 'data/raw/Respostas_Fomulario.xlsx'

respostas = Planilha.ler_formulario( forms_filepath )

# Colocar id

# Cliente.update_planilha_clientes( clientes, c_filepath )
# Vendedor.update_planilha_vendedor( vendedor, v_filepath )

# Cliente.gerar_contratos( clientes )

print( respostas[0] )

