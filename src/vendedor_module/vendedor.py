import os
from openpyxl import Workbook, load_workbook
from planilha_module import Planilha

class Vendedor:
    # atributos da classe
    def __init__(self, planilha: Planilha):
        self.nome = planilha.vendedor_nome 
        self.email = planilha.vendedor_email
        self.data_venda = planilha.venda_data
        self.valor_venda = planilha.venda_valor_total
        self.comissao = self.calcular_comissao()

    def calcular_comissao(self):
        nomeVendedorDoEmail = self.email.split('@')[0]

        # comissao dos vendedores
        vendedoresAtivos = {
            'marconnirodrigues': 100,
            'katianeves': 40,
            'thiagogiardini': 40,
            'wagneroliveira': 30,
        }
    
        # pega o vendedor da lista, se não tiver presente ele recebe 0% de comissão
        percentualDeComissao = vendedoresAtivos.get(nomeVendedorDoEmail, 0)

        self.comissao = (percentualDeComissao / 100) * self.valor_venda

    # to string
    def valor_comissao(self):
        return f'O vendedor {self.nome} vai receber R${self.comissao:.2f} de comissão.'

def atualizarPlanilha(vendedores, nomeDoArquivo="Vendedores.xlsx"):
    caminho = os.path.join("data", "processed", nomeDoArquivo)
    
    # Criar diretório se não existir
    os.makedirs("data/processed", exist_ok=True)
    
    # Verifica se o arquivo já existe para carregar ou criar um novo
    if os.path.exists(caminho):
        workbook = load_workbook(caminho)
    else:
        workbook = Workbook()

    for vendedor in vendedores:
        semana_venda = vendedor.data_venda.isocalendar()[1]
        nome_planilha = f"{vendedor.nome} - Semana {semana_venda}"

        # Se a planilha do vendedor para a semana não existir, cria uma nova
        if nome_planilha not in workbook.sheetnames:
            sheet = workbook.create_sheet(title=nome_planilha)
            # Adicionar cabeçalhos
            sheet.append(["Nome do Vendedor", "Total Vendido", "Comissão"])
        else:
            sheet = workbook[nome_planilha]
    
        # Verificar se há mudança de semana e adicionar linha de "X" se necessário
        if sheet.max_row > 1:
            ultima_semana = sheet.cell(row=sheet.max_row, column=1).value
            if ultima_semana != semana_venda:
                sheet.append(["X"] * 3)  # Adiciona uma linha com "X"
    
        # Calcular a comissão
        vendedor.calcular_comissao()
    
        # Preencher a planilha com dados
        sheet.append([vendedor.nome, vendedor.valor_venda, vendedor.comissao])

    # Remover a planilha padrão criada automaticamente se ela existir e for vazia
    if "Sheet" in workbook.sheetnames:
        del workbook["Sheet"]
    
    # Salvar a planilha
    workbook.save(caminho)
    print("Planilha VENDEDOR criada com sucesso!")
# end atualizarPlanilha ( )