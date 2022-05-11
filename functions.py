
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from imap_tools import AND, MailBox
import email.mime.application
import email.message
import os

# Capturar as informações do email modelo
def attachments(login, pwd, from_email, email_subject, att_name):
    with MailBox('imap.gmail.com').login(login, pwd) as mailbox:
        for message in mailbox.fetch(AND(from_=from_email)):
            if email_subject in message.subject: 
                content = message.html
                subject = message.subject
                print('Email encontrado')
                #se tiver anexo
                if len(message.attachments) > 0:
                    for att in message.attachments:
                        #verificando nome
                        if att_name in att.filename:
                            #transformar os dados do anexo em bytes
                            bytes_att = att.payload
                            #escrever o arquivo a partir dos bytes
                            with open(att.filename, 'wb') as file_excel:
                                file_excel.write(bytes_att)

                            path = f'.\\{att.filename}'
                            attachment = open(path,'rb')
                            att = email.mime.application.MIMEApplication(attachment.read(),name=os.path.basename(path))
                            attachment.close()  
                            return (att, content, subject)
                        else:
                            bytes_att = att.payload
                            #escrever o arquivo a partir dos bytes
                            with open(att.filename, 'wb') as file_excel:
                                file_excel.write(bytes_att)

                            path = f'.\\{att.filename}'
                            attachment = open(path,'rb')
                            att = email.mime.application.MIMEApplication(attachment.read(),name=os.path.basename(path))
                            attachment.close()  
                            return (att, content, subject)
                else:
                    print('Não há anexos')
                    return (content, subject)
            else:
                print('Email não encontrado')

# Enviar o email com o conteúdo capturado
def send_content(login, s, item, subject, content, att):
    msg = MIMEMultipart()
    body = MIMEMultipart()
    name = item['first_name']
    msg['Subject'] = subject
    msg['From'] = login
    msg.add_header('Content-Type', 'text/html')
    body.attach(MIMEText(content,'html'))
    msg.attach((body))
    if att:
        att.add_header('Content-Disposition', 'attachment')
        msg.attach(att)  
    s.sendmail(msg['From'], item['email'], msg.as_string().encode('utf-8'))
    print('enviado para:', item['email'])
