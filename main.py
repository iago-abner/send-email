from time import sleep
import smtplib
import random
import json

from functions import attachments, send_content

# informações p/ logar no gmail
login = 'iagoabner1@gmail.com'
password = 'Zzenilde'

# informações de pesquisa
email_subject = 'sabia que você pode acumular descontos com o CONTHABIL?'
from_email = 'iagoabner.ia@gmail.com'

# # Descomentar caso tenha anexo
# att_name = 'planilhateste'

# Capturando dados do email recebido
attachment = attachments(login = 'iagoabner1@gmail.com',
pwd = 'Zzenilde',
from_email = from_email,
email_subject = email_subject,
att_name = 'planilhateste')

if len(attachment) != 2:
    att, content, subject = attachment
elif len(attachment) == 2:
    content, subject = attachment
    att = False

# Conexão com o servidor
s = smtplib.SMTP('smtp.gmail.com:587')
s.ehlo()
s.starttls()
s.login(login, password)

# Leitura json
data_email = None
with open('data.json') as json_file:
    data_email = json.load(json_file)

# Envio dos emails
for i, item in enumerate(data_email['itens']):
    send_content(login, s, item, subject, content, att)
    # time = random.randint(5,60)
    # sleep(time)
    
print('Envios finalizados')
s.quit()