import pandas as pd

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a  base de dados
pd.set_option('display.max_columns',None) #Mostrar o máximo de colunas
print(tabela_vendas)

#faturamento por loja
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
#Voce está pegando o valor do faturamento e somando com todas as lojas ****
print(faturamento)

#quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)

#tickete médio por produto em cada loja
#Basicamente voce está pegando o valor final do faturamente e dividindo pela quantidade da tabela quantidade e este
#to_frame -- transformara tudi isso em uma tabela para nós !!!
ticket_medio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()


# enviar um  email com o relatorio
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = ''
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>" "</p>
'''

mail.Send()
