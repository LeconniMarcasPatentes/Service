import pandas as pd
from docx import Document
from datetime import datetime

# Classe Cliente
class Cliente:
    def __init__(self, razao_social, pf_pj, site_empresa, cpf_cnpj, tipo_empresa, 
                 endereco, telefones_empresa, nome_contato, email_contato, cargo_contato, 
                 telefones_contato, apresentacao):
        self.razao_social      = razao_social       # Razão Social/Nome
        self.pf_pj             = pf_pj              # PF/PJ
        self.site_empresa      = site_empresa       # Endereço do Site da Empresa
        self.cpf_cnpj          = cpf_cnpj           # CPF/CNPJ
        self.tipo_empresa      = tipo_empresa       # Tipo da Empresa (PJ, MEI, etc.)
        self.endereco          = endereco           # Endereço da Empresa (Dicionário)
        self.telefones_empresa = telefones_empresa  # Telefones da Empresa (Lista)
        self.nome_contato      = nome_contato       # Nome do Contato
        self.email_contato     = email_contato      # E-mail do Contato
        self.cargo_contato     = cargo_contato      # Cargo do Contato
        self.telefones_contato = telefones_contato  # Telefones de Contatos (Lista)
        self.apresentacao      = apresentacao       # Tipo de Apresentação (Figurativa, Mista, Nominativa)
        
    def __repr__(self):
        return f"Cliente(Razão Social: {self.razao_social}, CPF/CNPJ: {self.cpf_cnpj}, Endereço: {self.endereco}, Porte: {self.tipo_empresa})"
# end class Cliente

# Função para ler os dados da planilha e criar objetos Cliente
def read_from_excel ( filepath ):
    # Ler planilha do Excel
    df = pd.read_excel( filepath )

    # Lista de Cliente
    clientes = []

    # Iterar sobre cada linha da planilha e criar um objeto Cliente
    for index, row in df.iterrows( ):
        
        # Construir endereço como um dicionário
        endereco = {
            "logradouro" : row['LOGRADOURO'],
            "numero"     : row['NUMERO'],
            "bairro"     : row['BAIRRO/DISTRITO'],
            "complemento": row['COMPLEMENTO'],
            "municipio"  : row['MUNICIPIO'],
            "uf"         : row['UF'],
            "cep"        : row['CEP']
        }

        # Capturar telefones da empresa
        telefones_empresa = [row['TELEFONE'], row.get('TELEFONE 2', None), row.get('TELEFONE 3', None)]
        telefones_empresa = [tel for tel in telefones_empresa if tel]  # Remover valores nulos

        # Capturar telefones de contato
        telefones_contato = [row['TELEFONE/CELULAR'], row.get('TELEFONE/CELULAR 2', None), row.get('TELEFONE/CELULAR 3', None)]
        telefones_contato = [tel for tel in telefones_contato if tel]  # Remover valores nulos

        # Capturar CPF ou CNPJ dependendo do tipo de pessoa (PF ou PJ)
        cpf_cnpj = row['CPF'] if row['PF/PJ'] == 'PF' else row['CNPJ']

        # Criar o objeto Cliente
        cliente = Cliente(
            razao_social      = row['RAZAO SOCIAL'],
            pf_pj             = row['PF/PJ'],
            site_empresa      = row['SITE'],
            cpf_cnpj          = cpf_cnpj,
            tipo_empresa      = row['TIPO'],
            endereco          = endereco,
            telefones_empresa = telefones_empresa,
            nome_contato      = row['NOME'],
            email_contato     = row['E-MAIL'],
            cargo_contato     = row['CARGO'],
            telefones_contato = telefones_contato,
            apresentacao      = row['APRESENTACAO']
        )

        # Adicionar o novo cliente à lista
        clientes.append( cliente )

    return clientes  # Retornar lista de Clientes
pass
# end read_from_excel ( )

# Função para salvar os clientes em uma planilha Excel
def write_to_excel( clientes, output_filepath ):
    # Lista para armazenar os dados dos clientes como dicionários
    data = []

    # Iterar sobre os clientes e criar um dicionário com os dados
    for cliente in clientes:
        cliente_data = {
            'Razao Social'      : cliente.razao_social,
            'PF/PJ'             : cliente.pf_pj,
            'CPF/CNPJ'          : cliente.cpf_cnpj,
            'Tipo Empresa'      : cliente.tipo_empresa,
            'Logradouro'        : cliente.endereco['logradouro'],
            'Numero'            : cliente.endereco['numero'],
            'Bairro'            : cliente.endereco['bairro'],
            'Complemento'       : cliente.endereco['complemento'],
            'Municipio'         : cliente.endereco['municipio'],
            'UF'                : cliente.endereco['uf'],
            'CEP'               : cliente.endereco['cep'],
            'Telefone Empresa 1': cliente.telefones_empresa[0] if len(cliente.telefones_empresa) > 0 else '',
            'Telefone Empresa 2': cliente.telefones_empresa[1] if len(cliente.telefones_empresa) > 1 else '',
            'Telefone Empresa 3': cliente.telefones_empresa[2] if len(cliente.telefones_empresa) > 1 else '',
            'Site Empresa'      : cliente.site_empresa,
            'Nome Contato'      : cliente.nome_contato,
            'Email Contato'     : cliente.email_contato,
            'Cargo Contato'     : cliente.cargo_contato,
            'Telefone Contato 1': cliente.telefones_contato[0] if len(cliente.telefones_contato) > 0 else '',
            'Telefone Contato 2': cliente.telefones_contato[1] if len(cliente.telefones_contato) > 1 else '',
            'Telefone Contato 3': cliente.telefones_contato[2] if len(cliente.telefones_contato) > 1 else '',
            'Apresentacao'      : cliente.apresentacao
        }
        
        # Adicionar o dicionário à lista de dados
        data.append( cliente_data )

    # Converter a lista de dicionários em um DataFrame do pandas
    df = pd.DataFrame( data )

    # Salvar o DataFrame em um arquivo Excel
    df.to_excel( output_filepath, index=False )
pass
# end write_to_excel ( )

def write_to_docs ( template_path, output_path, cliente ):
    doc = Document( template_path )
    
    data_atual = datetime.now().strftime("%d/%m/%Y")  # Formato: dia/mês/ano
    
    for paragrafo in doc.paragraphs :
        texto = paragrafo.text
        texto = texto.split()
        
        if '[NOME_CONTRATANTE],' in texto:
            paragrafo.text = paragrafo.text.replace( '[NOME_CONTRATANTE],', cliente.razao_social.upper( )+ "," )
            
        if '[NACIONALIDADE_CONTRATANTE]' in texto:
            paragrafo.text = paragrafo.text.replace( '[NACIONALIDADE_CONTRATANTE]', 'brasileira' )
            
        if '[CNPJ_CONTRATANTE],' in texto:
            paragrafo.text = paragrafo.text.replace( '[CNPJ_CONTRATANTE],', formatar_cpf(cliente.cpf_cnpj)+"," if cliente.pf_pj == 'PF' 
                                                    else formatar_cnpj(cliente.cpf_cnpj)+"," ) 
                
        if '[ENDERECO_CONTRATANTE],' in texto:
            endereco_completo = f"{cliente.endereco['logradouro']}, {cliente.endereco['numero']}, {cliente.endereco['bairro']}, {cliente.endereco['municipio']}/{cliente.endereco['uf']} - {formatar_cep(cliente.endereco['cep'])}"  
            paragrafo.text = paragrafo.text.replace('[ENDERECO_CONTRATANTE],', endereco_completo+",")
                        
        if '[TITULO_PATENTE]' in texto:
            paragrafo.text = paragrafo.text.replace( '[TITULO_PATENTE]', cliente.razao_social.upper( ) )   
        
        if '[VALOR_SERVIÇO]' in texto:
            paragrafo.text = paragrafo.text.replace( '[VALOR_SERVIÇO]', '10.000,00')  
                
        if '[VALOR_INICIAL]' in texto:
            paragrafo.text = paragrafo.text.replace( '[VALOR_INICIAL]', '5.000,00') 
            
        if '[VALOR_FINAL]' in texto:
            paragrafo.text = paragrafo.text.replace( '[VALOR_FINAL]', '5.000,00')  
                
        if '[CIDADE_CONTRATANTE]' in texto:
            paragrafo.text = paragrafo.text.replace( '[CIDADE_CONTRATANTE]', cliente.endereco['municipio'] )
                    
        if '[LOCAL_DATA]' in texto:
            paragrafo.text = paragrafo.text.replace( '[LOCAL_DATA]', data_atual )

    doc.save( output_path )
pass
# end write_to_docs ( )

# Formatar CPF no padrão 000.000.000-00
def formatar_cpf ( cpf ):
    cpf = str(cpf).zfill(11)
    return f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}'
# end formatar_cpf

# Formatar o CNPJ no padrão 00.000.000/0000-00
def formatar_cnpj ( cnpj ):
    cnpj = str(cnpj).zfill(14)
    return f'{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}'
# end formatar_cnpj

# Formatar o CEP no padrão 00000-000
def formatar_cep ( cep ):
    cep = str(cep).zfill(8)
    return f'{cep[:5]}-{cep[5:]}'
# end formatar_cep
