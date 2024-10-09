import pandas as pd
from openpyxl import Workbook, load_workbook

class Vendedor:
    def __init__(self, vendedor_nome, vendedor_email, venda_data, venda_valor_total):
        self.vendedor_nome = vendedor_nome 
        self.vendedor_email = vendedor_email
        self.venda_data = venda_data
        self.venda_valor_total = venda_valor_total
        self.vendedor_comissao = 0.0
    # __init__ ( )

    """ Vendedor toString """
    def __str__(self) -> str:
        return (f"Vendedor:\n"
                f"Nome.................: {self.vendedor_nome    }\n"
                f"E-mail...............: {self.vendedor_email   }\n"
                f"Data da Venda........: {self.venda_data       }\n"
                f"Valor Total da Venda.: {self.venda_valor_total}\n"
                f"Valor da Comissão....: {self.vendedor_comissao}\n"
                )
    # __str__ ( )
    
    """ Atualizar a planilha de vendedores, adicionando os valores semanalmente. """
    def gravar_planilha_vendedores( vendedores ):
        output_filepath = 'data/processed/Vendedores.xlsx'
        
        try:
            try:
                df_existente = pd.read_excel(output_filepath)
            except FileNotFoundError:
                df_existente = pd.DataFrame()
            
            # Verificar se a coluna 'NOME INPI' existe, caso contrário, criar uma nova
            if 'NOME INPI' not in df_existente.columns:
                df_existente['NOME INPI'] = None

            novos_vendedores = []

            for res in vendedores:
                vendedor_comissao = calcular_comissao( res )
                
                vendedores_data = {
                    'NOME'                 : res.vendedor_nome,
                    'EMAIL'                : res.vendedor_email,
                    'DATA DA VENDA'        : pd.to_datetime(res.venda_data).strftime('%d/%m/%Y'),
                    'VALOR TOTAL DA VENDA' : res.venda_valor_total,
                    'COMISSAO'             : vendedor_comissao,
                    'NOME INPI'            : res.cliente_nome_inpi
                }
                
                venda_existente = df_existente[
                    (df_existente['NOME INPI'] == vendedores_data['NOME INPI']) 
                ]
                
                if venda_existente.empty:
                    novos_vendedores.append(vendedores_data)
                else:
                    df_existente.update(pd.DataFrame([vendedores_data]))

            df_novos = pd.DataFrame(novos_vendedores)
            df_final = pd.concat([df_existente, df_novos], ignore_index=True)

            df_final.to_excel(output_filepath, index=False)

            print(f"Planilha de vendedores atualizada com sucesso: {output_filepath}")

        except Exception as e:
            print(f"Erro ao atualizar a planilha de vendedores: {e}")
    # end gravar_planilha_vendedores ( )
# Vendedor

""" Calcular a comissão com base no vendedor """
def calcular_comissao( vendedor ):
    nomeVendedorDoEmail = vendedor.vendedor_email.split('@')[0]

    # Comissão dos vendedores
    vendedoresAtivos = {
        'marconnirodrigues': 100,
        'katianeves': 40,
        'thiagogiardini': 40,
        'wagneroliveira': 30,
    }

    percentualDeComissao = vendedoresAtivos.get(nomeVendedorDoEmail, 0)
    vendedor_comissao = (percentualDeComissao / 100) * vendedor.venda_valor_total
    return vendedor_comissao
# end calcular_comissao ( )