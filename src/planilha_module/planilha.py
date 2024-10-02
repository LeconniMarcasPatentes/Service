import json
import re
import pandas as pd

class Planilha:
    def __init__(self, vendedor_nome, venda_data, vendedor_email, cliente_razao_social, 
                 cliente_pf_pj, cliente_site, cliente_cpf_cpnj, cliente_tipo, cliente_endereco, cliente_telefone,
                 contato_nome, contato_email, contato_cargo, contato_telefone, venda_valor_total, venda_valor_inicial,
                 venda_valor_final, venda_servico, venda_pagante, venda_tipo, venda_recebiemento, venda_qtd_vezes, 
                 venda_prestacao, cliente_apresentacao, cliente_nome_inpi):
        self.vendedor_nome        = vendedor_nome
        self.venda_data           = venda_data
        self.vendedor_email       = vendedor_email
        self.cliente_razao_social = cliente_razao_social
        self.cliente_pf_pj        = cliente_pf_pj
        self.cliente_site         = cliente_site
        self.cliente_cpf_cpnj     = cliente_cpf_cpnj
        self.cliente_tipo         = cliente_tipo
        self.cliente_endereco     = cliente_endereco
        self.cliente_telefone     = cliente_telefone 
        self.contato_nome         = contato_nome
        self.contato_email        = contato_email
        self.contato_cargo        = contato_cargo
        self.contato_telefone     = contato_telefone 
        self.venda_valor_total    = venda_valor_total
        self.venda_valor_inicial  = venda_valor_inicial
        self.venda_valor_final    = venda_valor_final
        self.venda_servico        = venda_servico
        self.venda_pagante        = venda_pagante
        self.venda_tipo           = venda_tipo
        self.venda_recebiemento   = venda_recebiemento
        self.venda_qtd_vezes      = venda_qtd_vezes
        self.venda_prestacao      = venda_prestacao
        self.cliente_apresentacao = cliente_apresentacao
        self.cliente_nome_inpi    = cliente_nome_inpi
    # __init__ ( )
    
    """ toString """
    def __str__(self):
        return (f"Resposta:\n"
                f"\tNome do Vendedor.........: {self.vendedor_nome}\n"
                f"\tData da Venda............: {self.venda_data}\n"
                f"\tEmail do Vendedor........: {self.vendedor_email}\n"
                f"\tRazão Social do Cliente..: {self.cliente_razao_social}\n"
                f"\tPessoa Física ou Jurídica: {self.cliente_pf_pj}\n"
                f"\tCPF ou CNPJ do Cliente...: {self.cliente_cpf_cpnj}\n"
                f"\tTipo do Cliente..........: {self.cliente_tipo}\n"
                f"\tEndereço do Cliente......: {self.cliente_endereco}\n"
                f"\tTelefone do Cliente......: {self.cliente_telefone}\n"
                f"\tNome no INPI.............: {self.cliente_nome_inpi}\n"
                f"\tSite do Cliente..........: {self.cliente_site}\n"
                f"\tNome do Contato..........: {self.contato_nome}\n"
                f"\tEmail do Contato.........: {self.contato_email}\n"
                f"\tCargo do Contato.........: {self.contato_cargo}\n"
                f"\tTelefone do Contato......: {self.contato_telefone}\n"
                f"\tValor Total da Venda.....: {self.venda_valor_total}\n"
                f"\tValor Inicial da Venda...: {self.venda_valor_inicial}\n"
                f"\tValor Final da Venda.....: {self.venda_valor_final}\n"
                f"\tServiço da Venda.........: {self.venda_servico}\n"
                f"\tPagante da Venda.........: {self.venda_pagante}\n"
                f"\tTipo da Venda............: {self.venda_tipo}\n"
                f"\tRecebimento da Venda.....: {self.venda_recebiemento}\n"
                f"\tQuantidade de Vezes......: {self.venda_qtd_vezes}\n"
                f"\tPrestação................: {self.venda_prestacao}\n"
                f"\tApresentação do Cliente..: {self.cliente_apresentacao}\n"
                )
    # __str__ ( )

    """ Getters e Setters dos atributos """
    # vendedor_nome
    @property
    def vendedor_nome(self):
        return self._vendedor_nome

    @vendedor_nome.setter
    def vendedor_nome(self, value):
        self._vendedor_nome = value

    # venda_data
    @property
    def venda_data(self):
        return self._venda_data

    @venda_data.setter
    def venda_data(self, value):
        self._venda_data = value

    # vendedor_email
    @property
    def vendedor_email(self):
        return self._vendedor_email

    @vendedor_email.setter
    def vendedor_email(self, value):
        if not re.match(r"[^@]+@[^@]+\.[^@]+", value):
            raise ValueError("E-mail inválido")
        self._vendedor_email = value

    # cliente_razao_social
    @property
    def cliente_razao_social(self):
        return self._cliente_razao_social

    @cliente_razao_social.setter
    def cliente_razao_social(self, value):
        self._cliente_razao_social = value

    # cliente_pf_pj
    @property
    def cliente_pf_pj(self):
        return self._cliente_pf_pj

    @cliente_pf_pj.setter
    def cliente_pf_pj(self, value):
        self._cliente_pf_pj = value

    # cliente_site
    @property
    def cliente_site(self):
        return self._cliente_site

    @cliente_site.setter
    def cliente_site(self, value):
        self._cliente_site = value

    # cliente_cpf_cpnj
    @property
    def cliente_cpf_cpnj(self):
        return self._cliente_cpf_cpnj

    @cliente_cpf_cpnj.setter
    def cliente_cpf_cpnj(self, value):
        self._cliente_cpf_cpnj = value

    # cliente_tipo
    @property
    def cliente_tipo(self):
        return self._cliente_tipo

    @cliente_tipo.setter
    def cliente_tipo(self, value):
        self._cliente_tipo = value

    # cliente_endereco
    @property
    def cliente_endereco(self):
        return self._cliente_endereco

    @cliente_endereco.setter
    def cliente_endereco(self, value):
        self._cliente_endereco = value

    # cliente_telefone
    @property
    def cliente_telefone(self):
        return self._cliente_telefone

    @cliente_telefone.setter
    def cliente_telefone(self, value):
        self._cliente_telefone = value

    # contato_nome
    @property
    def contato_nome(self):
        return self._contato_nome

    @contato_nome.setter
    def contato_nome(self, value):
        self._contato_nome = value

    # contato_email
    @property
    def contato_email(self):
        return self._contato_email

    @contato_email.setter
    def contato_email(self, value):
        self._contato_email = value

    # contato_cargo
    @property
    def contato_cargo(self):
        return self._contato_cargo

    @contato_cargo.setter
    def contato_cargo(self, value):
        self._contato_cargo = value

    # contato_telefone
    @property
    def contato_telefone(self):
        return self._contato_telefone

    @contato_telefone.setter
    def contato_telefone(self, value):
        self._contato_telefone = value

    # venda_valor_total
    @property
    def venda_valor_total(self):
        return self._venda_valor_total

    @venda_valor_total.setter
    def venda_valor_total(self, value):
        self._venda_valor_total = value

    # venda_valor_inicial
    @property
    def venda_valor_inicial(self):
        return self._venda_valor_inicial

    @venda_valor_inicial.setter
    def venda_valor_inicial(self, value):
        self._venda_valor_inicial = value

    # venda_valor_final
    @property
    def venda_valor_final(self):
        return self._venda_valor_final

    @venda_valor_final.setter
    def venda_valor_final(self, value):
        self._venda_valor_final = value

    # venda_servico
    @property
    def venda_servico(self):
        return self._venda_servico

    @venda_servico.setter
    def venda_servico(self, value):
        self._venda_servico = value

    # venda_pagante
    @property
    def venda_pagante(self):
        return self._venda_pagante

    @venda_pagante.setter
    def venda_pagante(self, value):
        self._venda_pagante = value

    # venda_tipo
    @property
    def venda_tipo(self):
        return self._venda_tipo

    @venda_tipo.setter
    def venda_tipo(self, value):
        self._venda_tipo = value

    # venda_recebiemento
    @property
    def venda_recebiemento(self):
        return self._venda_recebiemento

    @venda_recebiemento.setter
    def venda_recebiemento(self, value):
        self._venda_recebiemento = value

    # venda_qtd_vezes
    @property
    def venda_qtd_vezes(self):
        return self._venda_qtd_vezes

    @venda_qtd_vezes.setter
    def venda_qtd_vezes(self, value):
        self._venda_qtd_vezes = value

    # venda_prestacao
    @property
    def venda_prestacao(self):
        return self._venda_prestacao

    @venda_prestacao.setter
    def venda_prestacao(self, value):
        self._venda_prestacao = value

    # cliente_apresentacao
    @property
    def cliente_apresentacao(self):
        return self._cliente_apresentacao

    @cliente_apresentacao.setter
    def cliente_apresentacao(self, value):
        self._cliente_apresentacao = value

    # cliente_nome_inpi
    @property
    def cliente_nome_inpi(self):
        return self._cliente_nome_inpi

    @cliente_nome_inpi.setter
    def cliente_nome_inpi(self, value):
        self._cliente_nome_inpi = value
    
    """ Funoes auxiliares """    
    def formatar_cpf_cnpj(self):
        if len(self._cliente_cpf_cpnj) == 11:
            return f"{self._cliente_cpf_cpnj[:3]}.{self._cliente_cpf_cpnj[3:6]}.{self._cliente_cpf_cpnj[6:9]}-{self._cliente_cpf_cpnj[9:]}"
        elif len(self._cliente_cpf_cpnj) == 14:
            return f"{self._cliente_cpf_cpnj[:2]}.{self._cliente_cpf_cpnj[2:5]}.{self._cliente_cpf_cpnj[5:8]}/{self._cliente_cpf_cpnj[8:12]}-{self._cliente_cpf_cpnj[12:]}"
        return self._cliente_cpf_cpnj
    
    def to_json(self):
        return json.dumps(self.__dict__)
    
    def to_csv(self):
        return ','.join([str(value) for value in self.__dict__.values()])
    
    def clonar(self):
        return Planilha(**self.__dict__)
    
    def log_alteracao(self, atributo, valor_antigo, valor_novo):
        with open("alteracoes_log.txt", "a") as log:
            log.write(f"Atributo: {atributo}, De: {valor_antigo}, Para: {valor_novo}\n")

    def salvar_bd(self, conexao):
        cursor = conexao.cursor()
        query = "INSERT INTO planilhas (...) VALUES (...)"
        cursor.execute(query, (...))
        conexao.commit()

    """
        Ler planilha de respostas do formulário.
        @return - lista de respostas.
        @param  - caminho do arquivo.
    """
    def ler_formulario(filepath):
        
        df = pd.read_excel(filepath)

        respostas = []

        for index, row in df.iterrows():
            endereco = {
                "logradouro": row['LOGRADOURO'],
                "numero": row['NUMERO'],
                "bairro": row['BAIRRO/DISTRITO'],
                "complemento": row['COMPLEMENTO'],
                "municipio": row['MUNICIPIO'],
                "uf": row['UF'],
                "cep": row['CEP']
            }

            telefones_empresa = [row['TELEFONE'], row.get('TELEFONE 2', None), row.get('TELEFONE 3', None)]
            telefones_empresa = [tel for tel in telefones_empresa if tel]  # Remover valores nulos

            telefones_contato = [row['TELEFONE/CELULAR'], row.get('TELEFONE/CELULAR 2', None), row.get('TELEFONE/CELULAR 3', None)]
            telefones_contato = [tel for tel in telefones_contato if tel]  # Remover valores nulos

            cpf_cnpj = row['CPF'] if row['PF/PJ'] == 'PF' else row['CNPJ']

            resposta = Planilha(
                vendedor_nome        = row['VENDEDOR'],
                venda_data           = row['DATA DA VENDA'],
                vendedor_email       = row['E-MAIL VENDEDOR'],
                cliente_razao_social = row['RAZAO SOCIAL'],
                cliente_pf_pj        = row['PF/PJ'],
                cliente_site         = row['SITE'],
                cliente_cpf_cpnj     = cpf_cnpj,
                cliente_tipo         = row['TIPO'],
                cliente_endereco     = endereco,
                cliente_telefone     = telefones_empresa,
                contato_nome         = row['NOME CONTATO'],
                contato_email        = row['E-MAIL CONTATO'],
                contato_cargo        = row['CARGO'],
                contato_telefone     = telefones_contato,
                venda_valor_total    = row['VALOR TOTAL'],
                venda_valor_inicial  = row['VALOR INICIAL'],
                venda_valor_final    = row['VALOR FINAL'],
                venda_servico        = row['SERVICO'],
                venda_pagante        = row['PAGANTE'],
                venda_tipo           = row['TIPO DE PAGAMENTO'],
                venda_recebiemento   = row['RECEBIMENTO'],
                venda_qtd_vezes      = row['VEZES'],
                venda_prestacao      = row['PRESTACAO'],
                cliente_apresentacao = row['APRESENTACAO'],
                cliente_nome_inpi    = row['NOME INPI']
            )

            respostas.append(resposta)

        return respostas
    # ler_formulario ( )
    
# Planilha
