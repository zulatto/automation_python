#importando pandas para mecher com excel (pip install pandas)
import pandas as pd

#Importar base de dados (pip install pyopenxl)
tabela_vendas = pd.read_excel('Vendas.xlsx')

#Visualizar a base de dados
pd.set_option('display.max_columns', None)
#print(tabela_vendas)

#tratamento dos dados
#Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#print(faturamento)

#Qnt de produtos por loja
qtd = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
#print(qtd)

#Ticket médio por produto
ticket_medio = (faturamento['Valor Final']/qtd['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
#print(ticket_medio)

#Enviar email com relatorio, nescessário (pip install pywin32)
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'email'
mail.Subject = 'Subject'
mail.HTMLBody = f'''
<p>Example email body</p>
<p>Prezados,</p>
<p>Segue o relatório de vendas por loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>
{qtd.to_html()}

<p>Ticket Médio:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

'''
mail.Send()