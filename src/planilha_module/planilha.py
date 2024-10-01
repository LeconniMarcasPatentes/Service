import json
import re

# class Planilha
class Planilha:
    """_summary_
    """
    def __init__(self, formulario_data_hora, vendedor_nome, venda_data, vendedor_email, cliente_razao_social, 
                 cliente_pf_pj, cliente_site, cliente_cpf_cpnj, cliente_tipo, cliente_endereco, cliente_telefone,
                 contato_nome, contato_email, contato_cargo, contato_telefone, venda_valor_total, venda_valor_inicial,
                 venda_valor_final, venda_servico, venda_pagante, venda_tipo, venda_recebiemento, venda_qtd_vezes, 
                 venda_prestacao, cliente_apresentacao, cliente_nome_inpi):
        self.formulario_data_hora = formulario_data_hora
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

    # Getters and Setters
    
    # formulario_data_hora
    @property
    def formulario_data_hora(self):
        return self._formulario_data_hora

    @formulario_data_hora.setter
    def formulario_data_hora(self, value):
        self._formulario_data_hora = value

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
    
    # Outros / Planos Futuros
    
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
        
# Planilha