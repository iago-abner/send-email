from email.mime.multipart import MIMEMultipart
from email.mime.message import MIMEMessage
from email.mime.text import MIMEText
from imap_tools import MailBox, AND
from time import sleep
import email.mime.application
import email.message
import pandas as pd
import smtplib
import random
import json
import os

contatos = pd.read_csv('planilha.csv')

# informações p/ logar no gmail
login = 'iagoabner1@gmail.com'
pwd = 'Zzenilde'

# informações de pesquisa
email_subject = 'Você conhece o Robô CONTHABIL?'
from_email = ' inaldo@conthabilbr.com'
att_name = 'planilhateste'

# Capturando dados do email recebido
data = {'att': False}
with MailBox('imap.gmail.com').login(login, pwd) as mailbox:
    for msg in mailbox.fetch(AND(from_=from_email)):
        if email_subject in msg.subject:
            data.update({'content': msg.html})
            data.update({'subject': msg.subject}) 
            # se tiver anexo
            if len(msg.attachments) > 0:
                # percorra os anexos
                for att in msg.attachments:
                    # verificando nome
                    if att_name in att.filename:
                        # transformar os dados do anexo em bytes
                        bytes_att = att.payload
                        # escrever o arquivo a partir dos bytes
                        with open(att.filename, 'wb') as file_excel:
                            file_excel.write(bytes_att)

                        path = f'.\\{att.filename}'
                        attachment = open(path, 'rb')
                        att = email.mime.application.MIMEApplication(attachment.read(), name=os.path.basename(path))
                        attachment.close()

                        data.update({'att': att})
                break
            else:
                print('Não há anexos')
                break
        else:
            print('Email não encontrado')

# Conectar com o servidor
s = smtplib.SMTP('smtp.gmail.com:587')
s.ehlo()
s.starttls()

# Login do email
s.login(login, pwd)

# Leitura json
data_email = None
with open('data.json') as json_file:
    data_email = json.load(json_file)

# Envio do email
for i, item in enumerate(data_email['itens']):
    msg = MIMEMultipart()
    body = MIMEMultipart()
    name = item['first_name']
    msg['Subject'] = data['subject']
    msg['From'] = login
    msg.add_header('Content-Type', 'text/html')
    body.attach(MIMEText(f'Oi, {name}. Tudo tranquilo?', 'html'))
    body.attach(MIMEText(data['content'],'html'))
    msg.attach((body))
    if data['att']:
        data['att'].add_header('Content-Disposition', 'attachment')
        msg.attach(data['att'])  
    s.sendmail(msg['From'], item['email'], msg.as_string().encode('utf-8'))
    print('enviado para:', item['email'])
    time = random.randint(5,60)
    sleep(time)
print('Envios finalizados')
s.quit()