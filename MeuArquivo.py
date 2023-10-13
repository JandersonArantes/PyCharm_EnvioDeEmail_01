import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados (tratamento)
# pd.set_option('display.max_columns', None)
# print(tabela_vendas)
# print(tabela_vendas[['ID Loja', 'Valor Final']])

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', "Valor Final"]].groupby('ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# Ticket médio por produto em cada loja
print('-' * 50)
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Enviar um relatório por e-mail
outlook = win32.Dispatch('Outlook.Application')
mail = outlook.CreateItem(0)
mail.To = 'janderson_alves@hotmail.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o relatório Vendas Por Loja.</P>

<p>Faturamento:</P>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</P>
{quantidade.to_html()}

<p>Ticket Médio Por Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, estou à disposição.</p>
<p>At.te,</P>
<p>Janderson Alves De Arantes</P>
<p>Analista De Sistemas</P>
<p>(19) 99158-3325</p>
<p>Porto Ferreira - SP</p>
'''
mail.Send()
# mail.Display()
# Obs.: O  comando mail.Send() envia o e-mail.
# O comando mail.Display() abre o Outlook antes de enviar o e-mail.
print('E-mail enviado!')
