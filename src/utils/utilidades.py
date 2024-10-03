import json
import pandas as pd

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

# Formatar o CEP no padrão 00.000-000
def formatar_cep ( cep ):
    cep = str(cep).zfill(8)
    return f'{cep[:2]}.{cep[:5]}-{cep[5:]}'
# end formatar_cep

def formatar_cpf_cnpj( cpf_cnpj ):
    if len(cpf_cnpj) == 11:
        return f"{cpf_cnpj[:3]}.{cpf_cnpj[3:6]}.{cpf_cnpj[6:9]}-{cpf_cnpj[9:]}"
    elif len(cpf_cnpj) == 14:
        return f"{cpf_cnpj[:2]}.{cpf_cnpj[2:5]}.{cpf_cnpj[5:8]}/{cpf_cnpj[8:12]}-{cpf_cnpj[12:]}"
    return cpf_cnpj

def to_json( obj ):
    return json.dumps(obj.__dict__)

def to_csv( obj ):
    return ','.join([str(value) for value in obj.__dict__.values()])

def get_ultimo_id_planilha( filepath ):
    planilha = pd.read_excel( filepath )
    return planilha.dropna(how='all').shape[0]
