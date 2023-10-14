``` #Python001
Nesta aplicação, são realizados os procedimentos:
* Importação de base de dados
- Arquivo Excel Vendas.xlsx com 101.000 linhas
- Esta planilha possui as vendas de várias lojas
* Visualização da 'Tabela_Vendas' no terminal
* Criação e visualização da tabela 'Faturamento Por Loja' no terminal
* Criação e visualização da tabela 'Ticket Médio Por Loja" no terminal
* Envio de e-mail com os relatórios:
- Faturamento Por Loja
- Quantidade Total De Produtos Vendidos Por Loja
- Ticket Médio De Vendas Por Loja



*** Explicando o código ***
# No terminal, executar as instalações abaixo para poder manipular
# os dados do Excel
pip install pandas

# No terminal, instalar o pacote openpyxl para abrir planilha do Excel
pip install openpyxl

# No início do código importar e apelidar as bibliotecas necessárias
# para manipulação dos dados com o Excel
import pandas as pd

# Esta biblioteca é necessária para acessar o Outlook
import win32com.client as win32

# Importar a base de dados
# A variável 'tabela_vendas' recebe a tabela 'Vendas.xlsx'
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados (tratamento)
# Esta configuração determina a quantidade máxima de colunas para exibir.
# O argumento "None" determina que todas as colunas serão exibidas.
pd.set_option('display.max_columns', None)
# Exibe a 'tabela_vendas' com todas as colunas
print(tabela_vendas)

# Exibe apenas as colunas 'ID Loja' e 'Valor Final' da 'tabela_vendas'
print(tabela_vendas[['ID Loja', 'Valor Final']])

# Faturamento por loja
# A variável 'faturamento' recebe como resultado, uma tabela que é o a soma das
# vendas por loja.
faturamento = tabela_vendas[['ID Loja', "Valor Final"]].groupby('ID Loja').sum()

# Exibe a tabela 'faturamento'
print(faturamento)

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Exibe a tabela 'quantidade'
print(quantidade)

# Exibe 50 vezes o caractere -
print('-' * 50)

# Ticket médio por produto em cada loja
# O resultado da desta operação é do tipo float. Para converter o resultado em uma tabela,
# é utilizado o método to_frame()
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

# Alterando o nome '0' da coluna para o nome 'Ticket Médio'
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# Imprime a tabela 'ticket_medio'
print(ticket_medio)

# Enviar um relatório por e-mail
# A variável 'outlook' aponta para a aplicação 'Outlook'
outlook = win32.Dispatch('Outlook.Application')

# Inicia a criação de um e-mail
mail = outlook.CreateItem(0)

# Enviar e-mail para
mail.To = 'janderson_alves@hotmail.com'

# Assunto do e-mail
mail.Subject = 'Relatório de vendas por loja'

# Corpo do e-mail. A letra f na frente de ''', indica que o que estiver entre chaves é uma expressão e não um texto.
# O método to_html() converte a tabela em html
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
# Envia o e-mail
mail.Send()

# Abre o Outlook para editar o e-mail e depois enviar.
# mail.Display()

print('E-mail enviado!')
```
