from email import charset
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from imap_tools import MailBox, AND
import email.mime.application
import email.message
import smtplib
import os

#informações p/ logar no gmail
login = 'iagoabner1@gmail.com'
pwd = 'Zzenilde'

#informações de pesquisa
email_subject = 'assunto1'
from_email = 'iagoabner.ss@gmail.com'
att_name = 'planilhateste'

#informações de envio
email_to = 'iagoabner.ia@gmail.com'
subject_to = 'resposta ao envio'

#Capturando dados do email recebido
with MailBox('imap.gmail.com').login(login, pwd) as mailbox:
    for msg in mailbox.fetch(AND(subject=email_subject, from_=from_email)):
        content = msg.html
        #se tiver anexo
        if msg.attachments:
            #armazenando o conteúdo do texto 
            content = msg.html
            #percorra os anexos
            for att in msg.attachments:
                #verificando nome
                if att_name in att.filename:
                    #transformar os dados do anexo em bytes
                    bytes_att = att.payload
                    #escrever o arquivo a partir dos bytes
                    with open(att.filename, 'wb') as file_excel:
                        file_excel.write(bytes_att)

#Reenviando o email
msg = MIMEMultipart()
msg['Subject'] = subject_to
msg['From'] = login
msg['To'] = email_to
password = pwd
msg.add_header('Content-Type', 'text/html')
msg.attach(MIMEText(content, 'html'))

#Envio do anexo
path = '.\\planilhateste.xlsx'
attachment = open(path,'rb')

att = email.mime.application.MIMEApplication(attachment.read(), subtype="xslx", name=os.path.basename(path))
attachment.close()

att.add_header('Content-Disposition', 'attachment')
msg.attach(att)

s = smtplib.SMTP('smtp.gmail.com:587')
s.ehlo()
s.starttls()

#Credenciais para envio do email
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email enviado')
s.quit()