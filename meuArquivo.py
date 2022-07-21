# Importar a Biblioteca Panda

import pandas as pd
''' Pandas é uma biblioteca usada associar um projecto ou arquivo python com o excel.
fazendo assim a interação entre o projecto e uma base de dados em excel. '''

# Importar a base de dados

tabela_venda = pd.read_excel('Vendas.xlsx')

''' Então para que o arquivo em excel seja lido é necessário instalar uma outra Biblioteca que faz
leitura/abertura do arquivo python no arquivo xl. '''

# Visualizar a base de dados
pd.set_option('display.max_columns', None)

# Faturamento por loja
faturamento = tabela_venda[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
            # Comando usado para fazer filtragem dos campos na base de dados

print(faturamento)

# Quantidades de produtos vendidos por cada loja
quantidade = tabela_venda[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print("\n")
print(quantidade)


# ticket médio por produto em cada loja

# Enviar um email por relatório

print("Ola, python and Panda")
