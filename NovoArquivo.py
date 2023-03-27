#!/usr/bin/env python
# coding: utf-8

# In[ ]:

import pandas as pd
import smtplib
import email.message

# importar a base de dados  
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a bd
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# faturamento

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# qtd de prod vendidos p loja

quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

# ticket médio por loja

print(' ' * 50)
ticketmedio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticketmedio = ticketmedio.rename(columns={0: 'Ticket médio'})
print(ticketmedio)


# envia o email
print(' ' * 50)
def enviar_email():
    corpo_email = f"""
    <p>Prezados, segue o relatório das vendas:</p> <br>
    <p>Faturamento:</p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
    <p>Quantidade:</p>
    {quantidade.to_html()}
    <p>Ticket médio:</p>
    {ticketmedio.to_html(formatters={'Ticket médio': 'R${:,.2f}'.format})}
    <p>Qualquer dúvida estou à disposição.</p>
    <p>Atenciosamente, Tiago</p>
    """

    msg = email.message.Message()
    msg['Subject'] = "Relatório"
    msg['From'] = 'DIGITE O REMETENTE'
    msg['To'] = 'DIGITE O DESTINATÁRIO'
    password = 'DIGITE A SENHA GERADA AUTOMATICAMENTE'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')


# In[ ]:


enviar_email()
