import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#Iniciar servidor smtp:
host = "smtp.gmail.com"
port = "587"
login = "email"
senha = "senha"

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