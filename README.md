# Automacao_de_processos

Etapas do Projeto
1 Importação de Arquivos e Bibliotecas
- Importação das bibliotecas necessárias.
- Carregamento das bases de dados das vendas, lojas e e-mails.

2 Criação de Tabelas por Loja e Definição do Dia do Indicador
- Combinação dos dados das vendas com os nomes das lojas.
- Criação de um dicionário contendo as vendas separadas por loja.
- Definição do último dia de venda como o dia do indicador.

3 Salvamento das Planilhas na Pasta de Backup
- Verificação e criação da pasta de backup.
- Salvamento das planilhas de vendas de cada loja na respectiva pasta de backup.

4 Cálculo de Indicadores para Cada Loja
- Definição de metas diárias e anuais para faturamento, quantidade de produtos vendidos e ticket médio.
- Cálculo desses indicadores para cada loja.
- Preparação e envio do e-mail com os resultados para os gerentes de loja.

5 Envio de E-mail para Diretoria com Ranking
- Cálculo do ranking das lojas em termos de faturamento diário e anual.
- Preparação e envio do e-mail com os rankings para a diretoria.

# Projeto de Automação de Processos - Indicadores de Lojas
Descrição do Projeto
Este projeto tem como objetivo automatizar o processo de cálculo e envio de indicadores de desempenho para gerentes de lojas e diretoria de uma rede de lojas de varejo. As principais etapas incluem a importação de dados, criação de tabelas de vendas por loja, cálculo de indicadores, salvamento de arquivos de backup, envio de e-mails com resultados e criação de rankings.

Estrutura do Projeto
1. Importação de Arquivos e Bibliotecas
No primeiro passo, importamos as bibliotecas necessárias e carregamos as bases de dados das vendas, lojas e e-mails:

python
Copiar código
import pandas as pd
import win32com.client as win32
import pathlib

emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas = pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')

display(emails)
display(lojas)
display(vendas)
2. Criação de Tabelas por Loja e Definição do Dia do Indicador
Combinamos os dados das vendas com os nomes das lojas e criamos um dicionário onde cada chave é o nome de uma loja e o valor é um DataFrame contendo as vendas dessa loja. Também definimos o último dia de venda como o dia do indicador.

python
Copiar código
vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)

dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja, :]
display(dicionario_lojas['Rio Mar Recife'])
display(dicionario_lojas['Shopping Vila Velha'])

dia_indicador = vendas['Data'].max()
print(dia_indicador)
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))
3. Salvamento das Planilhas na Pasta de Backup
Verificamos se a pasta de backup existe, e criamos subpastas para cada loja, se necessário. Salvamos as planilhas de vendas de cada loja na respectiva pasta de backup.

python
Copiar código
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

if not caminho_backup.exists():
    caminho_backup.mkdir()

for loja in dicionario_lojas:
    pasta_loja = caminho_backup / loja
    if not pasta_loja.exists():
        pasta_loja.mkdir()

    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = pasta_loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)
4. Cálculo de Indicadores para Cada Loja
Definimos metas diárias e anuais para faturamento, quantidade de produtos vendidos e ticket médio. Calculamos esses indicadores para cada loja e preparamos o e-mail com os resultados.

python
Copiar código
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]

    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    qtde_produtos_ano = len(vendas_loja['Produto'].unique())
    qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())
    ticket_medio_ano = vendas_loja.groupby('Código Venda').sum()['Valor Final'].mean()
    ticket_medio_dia = vendas_loja_dia.groupby('Código Venda').sum()['Valor Final'].mean()
    
    # Condicionais para a cor dos indicadores
    cor_fat_dia = 'green' if faturamento_dia >= meta_faturamento_dia else 'red'
    cor_fat_ano = 'green' if faturamento_ano >= meta_faturamento_ano else 'red'
    cor_qtde_dia = 'green' if qtde_produtos_dia >= meta_qtdeprodutos_dia else 'red'
    cor_qtde_ano = 'green' if qtde_produtos_ano >= meta_qtdeprodutos_ano else 'red'
    cor_ticket_dia = 'green' if ticket_medio_dia >= meta_ticketmedio_dia else 'red'
    cor_ticket_ano = 'green' if ticket_medio_ano >= meta_ticketmedio_ano else 'red'

    # Envio do e-mail
    outlook = win32.Dispatch('outlook.application')
    nome = emails.loc[emails['Loja'] == loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    
    mail.HTMLBody = f'''
    <p>Bom dia, {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>

    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center">{qtde_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>

    <p>Qualquer dúvida estou à disposição.</p>
    <p>Att., Lira</p>
    '''
    
    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))
    mail.Send()
    print(f'E-mail da Loja {loja} enviado')
5. Envio de E-mail para Diretoria com Ranking
Calculamos o ranking das lojas em termos de faturamento diário e anual e enviamos esses rankings para a diretoria.

python
Copiar código
faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)
faturamento_dia = vendas.loc[vendas['Data'] == dia_indicador, :].groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_dia.sort_values(by='Valor Final', ascending=False)

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'

corpo_email = f'''
<p>Bom dia, Diretoria</p>

<p>Ranking de lojas por faturamento:</p>

<p><strong>Faturamento Diário ({dia_indicador.day}/{dia_indicador.month})</strong></p>
'''

for i, loja in enumerate(faturamento_lojas_dia.index):
    corpo_email += f'<p>{i+1}º - {loja} - R${faturamento_lojas_dia.loc[loja, "Valor Final"]:.2f}</p>'

corpo_email += '''
<p><strong>Faturamento Anual</strong></p>
'''

for i, loja in enumerate(faturamento_lojas_ano.index):
    corpo_email += f'<p>{i+1}º - {loja} - R${faturamento_lojas_ano.loc[loja, "Valor Final"]:.2f}</p>'

corpo_email += '''
<p>Qualquer dúvida estou à disposição.</p>

<p>Att., Italo</p>
'''

mail.HTMLBody = corpo_email
mail.Send()
print('E-mail da Diretoria enviado')


Conclusão
Este projeto automatiza o processo de coleta, cálculo e envio de indicadores de desempenho, economizando tempo e reduzindo a possibilidade de erros humanos. Com a implementação deste script, espera-se que os gerentes das lojas e a diretoria possam tomar decisões mais informadas e rápidas baseadas em dados precisos e atualizados.
