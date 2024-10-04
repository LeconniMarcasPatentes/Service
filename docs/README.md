# Documentação do Programa de Processamento de Informações de Cliente

## Descrição Geral

Este programa foi desenvolvido para automatizar a geração e envio de contratos baseados em informações de clientes contidas em uma planilha. Ele divide os dados gerais em duas planilhas separadas: uma para o vendedor, com informações pertinentes ao vendedor, e outra para o cliente, contendo os dados necessários para o cadastro no INPI. Além disso, o programa automatiza a personalização de um documento Word com as informações do cliente e envia o contrato finalizado para o vendedor, que o repassa ao cliente.

### Funcionalidades Principais

1. **Leitura da Planilha de Entrada**  
   O programa inicia importando uma planilha com informações gerais de um cliente. Essa planilha contém:
   - Nome, email, e contato do vendedor
   - Informações do cliente, como dados para cadastro no INPI

2. **Divisão da Planilha em Duas**  
   A planilha original é dividida em duas novas planilhas:
   - **Planilha Vendedor**: Contém os dados do vendedor, incluindo nome, email, contato e comissão.
   - **Planilha Cliente**: Contém todas as informações do cliente necessárias para o cadastro no INPI.

3. **Modificação do Documento de Contrato**  
   O programa pega um documento Word de contrato padrão e o preenche com as informações do cliente. Essa etapa inclui:
   - Inserção do nome do cliente
   - Atualização de dados cadastrais
   - Personalização de termos de acordo com as necessidades do cliente

4. **Envio do Contrato por Email**  
   Após gerar e modificar o contrato, o programa envia um email para o vendedor contendo o documento final. O vendedor é responsável por encaminhar o contrato ao cliente.

### Fluxo do Programa

1. **Entrada**: O programa recebe uma planilha com os dados gerais.
2. **Processamento**:
   - Divide a planilha em duas novas planilhas: uma para o vendedor e outra para o cliente.
   - Modifica um documento Word com as informações específicas do cliente.
3. **Saída**: Gera as planilhas e o documento Word modificados.
4. **Envio por Email**: O contrato é enviado ao vendedor por email.

### Pré-requisitos

- **Bibliotecas Python**:
  - `pandas` para manipulação de planilhas Excel
  - `openpyxl` ou `xlsxwriter` para escrever nas planilhas
  - `python-docx` para modificar arquivos Word
  - `smtplib` ou uma biblioteca similar para envio de emails

- **Arquivos de Entrada**:
  - Planilha Excel com dados do cliente e vendedor
  - Documento Word padrão de contrato

### Exemplo de Uso

1. Execute o programa fornecendo a planilha de entrada.
2. O programa gera as duas planilhas (vendedor e cliente) e o documento Word de contrato.
3. O contrato é enviado automaticamente ao email do vendedor.
