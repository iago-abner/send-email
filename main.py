from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from imap_tools import MailBox, AND
import email.mime.application
import email.message
import smtplib
import pandas as pd
import os

contatos = pd.read_csv('planilha.csv')

#informações p/ logar no gmail
login = 'iagoabner1@gmail.com'
pwd = 'Zzenilde'

#informações de pesquisa
email_subject = 'Você conhece o Robô CONTHABIL?'
from_email = 'inaldo@conthabilbr.com'
att_name = 'planilhateste'

#informações de envio


#Capturando dados do email recebido
with MailBox('imap.gmail.com').login(login, pwd) as mailbox:
    for msg in mailbox.fetch(AND(from_=from_email)):
        if email_subject in msg.subject: 
            content = msg.html
            subject_to = msg.subject
            print('ACHOU O EMAIL')
            #se tiver anexo
            if len(msg.attachments) > 0:
                #armazenando o conteúdo do texto 
                #percorra os anexos
                for att in msg.attachments:
                    #verificando nome
                    if att_name in att.filename:
                        #transformar os dados do anexo em bytes
                        bytes_att = att.payload
                        #escrever o arquivo a partir dos bytes
                        with open(att.filename, 'wb') as file_excel:
                            file_excel.write(bytes_att)

                        msg = MIMEMultipart()
                        msg.add_header('Content-Type', 'text/html')
                        msg.attach(MIMEText(content, 'html'))

                        path = f'.\\{att.filename}'
                        attachment = open(path,'rb')
                        att = email.mime.application.MIMEApplication(attachment.read(),name=os.path.basename(path))
                        attachment.close()

                        att.add_header('Content-Disposition', 'attachment')
                        msg.attach(att)        
                break
            else:
                name = 'iago'
                msg = MIMEMultipart()
                msg.add_header('Content-Type', 'text/html')
                msg.attach(MIMEText(f'Olá {name}, tudo bem?', 'html'))
                msg.attach(MIMEText(content, 'html'))
                print('Attachment not exists')
                break
        else:
            print(msg.subject)
            print('Email not found')

#Reenviando o email
msg['From'] = login
msg['Subject'] = subject_to
password = pwd

s = smtplib.SMTP('smtp.gmail.com:587')
s.ehlo()
s.starttls()

#Credenciais para envio do email
s.login(msg['From'],password)
for i, contato in enumerate(contatos['contato']):
    s.sendmail(msg['From'], str(contato), msg.as_string().encode('utf-8'))
    print('enviado para:', contato)
print('Email enviado')
s.quit()