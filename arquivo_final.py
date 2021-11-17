
#Excel
import pandas as pd
from pandas.core.indexes.period import PeriodIndex

#Email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders



#Parte relacionada com a cópia do arquivo excel

#Le o arquivo
arquivo = pd.read_excel(r"arquivo_original.xlsx")

#Coleta as linhas que serão separadas
linhaI = int(input("digite a primeira linha: "))
linhaF = int(input("digite a ultima linha: "))


#Define as variaveis
a = []
e = linhaI-2
n = 0

#Separa as linhas em lista
while e <= linhaF-2:
    
    a.append(e)
    n = n+1
    e = e+1


#Define "selecionadas" como as linhas escolhidas
selecionadas = arquivo.loc[a]


#Copia as linhas escolhidas em outro arquivo
selecionadas.to_excel('arquivo_enviado.xlsx',index=False)


#Parte dedicada a enviar o arquivo final por email.

#Iniciar servidor smtp:
host = "smtp.gmail.com"
port = "587"
login = "excel.Botsend@gmail.com"
senha = "tw0508alt"

server = smtplib.SMTP(host,port)

server.ehlo()
server.starttls()

server.login(login,senha)

#Montar o email:
corpo = "imprimir"

email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = login #para testes
email_msg['Subject'] = "Email automático Excel"

    #Configurar Anexo
cam_arquivo = "C:\\Users\luize\\OneDrive\\Documentos\\GitHub\\emailexcelbot\\arquivo_enviado.xlsx"
attchment = open(cam_arquivo,'rb')


att = MIMEBase('aplication','octet-stream')
att.set_payload(attchment.read())
encoders.encode_base64(att)


att.add_header('Content-Disposition',f'attachment; filename=arquivo_enviado.xlsx')
attchment.close()

email_msg.attach(att)
    #Anexo Configurado


email_msg.attach(MIMEText(corpo,'plain'))

#Enviar o email

server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())

server.quit()


#fonte: https://www.youtube.com/watch?v=umvzsQLZYD4