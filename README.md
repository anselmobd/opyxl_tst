# opyxl_tst
O objetivo deste repositório é testar os recursos básicos da biblioteca openpyxl.

É desenvolvido um programa que acessa os dados em um arquivo ".xlsx", processa-os e grava o resultado em outro arquivo ".xlsx".

# Origem da planilha utilizada

A planilha ".xlsx" para utilizada neste teste é enconrada na web. Não há nenhum significado específico na escolha da planilha.

Busca feita no Portal Brasileiro de Dados Abertos
https://dados.gov.br

A planilha escolhida se encontra em:

Empresa Mato-Grossense de Tecnologia da Informação

https://dados.gov.br/dados/organizacoes/visualizar/empresa-mato-grossense-de-tecnologia-da-informacao

Bens Móveis e Imóveis

https://dados.gov.br/dados/conjuntos-dados/bens-moveis-e-imoveis

XLSX - Computadores e Periféricos 2024

Listagem de Computadores e Periféricos de inventario 2024. 

Nome original da planilha: "01.xlsx"

Renomeada para: "equipamentos.xlsx"

# Descrição do processamento

Colunas utilizadas da planilha original:
- "Descr. Sint.": Descrição do equipamento
- "Dt.Aquisicao"
- "Quantidade"
- "Tipo Ativo"

Pré-processamento:
- Cada equipamento aparece na listagem três vezes. Para evitar essa repetição, são consideradas apenas as linhas onde o "Tipo Ativo" é vazio.
- Baseado na descrição, é criada uma classe de equipamento. Utilizando uma heurística simples são separadas algumas classes, como "monitor" e "notebook". Equipamentos não enquadrados nas classes básicas, ficam na classe "diverso".

Planilha criada:
- Apresenta totais de equipamentos adquiridos em uma grade por ano (coluna) e classe (linha)
- Apresenta formatações simples como alinhamento e negrito
- Totaliza os dados por ano e por classe
