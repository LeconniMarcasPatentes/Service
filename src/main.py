from planilha_module import Planilha

forms_filepath = 'data/raw/Respostas_Fomulario.xlsx'

respostas = Planilha.ler_formulario( forms_filepath )

print( respostas[0] )

