# Importar a Biblioteca Panda

import pandas as pd
''' Pandas é uma biblioteca usada para associar um projecto ou arquivo python com o excel.
fazendo assim a interação entre o projecto e uma base de dados em excel. '''

# Importar a base de dados

import win32com.client as win32
''' pywin32 é uma biblioteca usado enviar relatórios por email via código python'''

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
print("-" *50)
print(quantidade)

# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
        # A função 'to_frame' é utilizado para tranformar os resultado de uma operação com colunas de uma tabela
        # realizadada por ela numa tabela, tirando o retorno dtype: float64

print("\n")
print("-" *50)
print("Ticket Médio de cada Loja \n")
print(ticket_medio)

# Enviar um email por relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = "gomilson.otenta@gmail.com"
mail.subject = "Relatório de Vendas por Lojas"
mail.HTMLBody = f'''
    <center><h2> Relatório de Venda </h2></center>
    
    <br/>
    <hr/>
    <h5> Faturamento por valor final de cada Loja </h5>
    {faturamento.to_html(formatters={'Valor Final': 'Kz{:,.2f}'.format})}
    
    <br/>
    <hr/>
    <h5> Quantidade de produto por cada Loja </h5>
    {quantidade.to_html()}
    
    <br/>
    <hr/>
    <h5> Ticket Médio por cada Loja </h5>
    {ticket_medio.to_html(formatters={'Ticket Médio': 'Kz{:,.2f}'.format})}
'''
# This field is optional

# To attach a file to email (optional)
#attachment = ""
#mail.attachments.add(attachment)

mail.Send()

print("Email enviado com sucesso!")
