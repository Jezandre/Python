import pandas as pd 
import datetime 
import smtplib 
import datetime 
import smtplib 
import datetime 
import os 
import xlrd
import pyodbc
from email.mime.text import MIMEText 
from email.mime.multipart import MIMEMultipart 

# Define a conexão com o SQL Server usando a autenticação do Windows
server = '<<NOME DO SERVIDOR>>'
database = '<<NOME DO BANCO>>'
cnxn = pyodbc.connect('Driver={SQL Server};'
                      'Server=' + server + ';'
                      'Database=' + database + ';'
                      'Trusted_Connection=yes;')

# Executa uma consulta SQL

query = ("<<QUERY>>")

df = pd.read_sql(query, cnxn)
print(df.head(1000))

# Define o servidor SMTP e porta 
smtp_server = "outlook.office365.com" 
smtp_port = 587  

# Informações da conta de e-mail que enviará as mensagens 
sender_email = "<<E-MAIL>>" 
sender_password = "<<SENHA>>"  

# Percorre cada linha da tabela 
for i in range(len(df)): 
    # Obtém a data de vencimento da linha
    
    vencimento = df.iloc[i]["<<NOME DA COLUNA>>"]
    if vencimento is not None:
        vencimento = datetime.datetime.strptime(vencimento, '%Y-%m-%d')
    # Verifica se a data de vencimento é igual ou anterior à data atual 

    if vencimento is not None and vencimento <= datetime.datetime.now(): 
        # Se sim, obtém o e-mail do responsável e envia a mensagem 
        msg = MIMEMultipart() 
        responsavel = df.iloc[i]["USUARIO"] 
        tarefa = df.iloc[i]["NOME"] 
        recipient_email = df.iloc[i]["EMAIL"] 
        msg['Subject'] = f"Alerta de documento vencido - {tarefa}" 
        body = f"Ola {responsavel},\n\nO prazo para a revisão do documento {tarefa} venceu no dia {vencimento}.\n\nAtenciosamente,\nSua equipe de trabalho" 
        msg.attach(MIMEText(body, 'plain'))               

        # Conecta ao servidor SMTP e envia a mensagem 
        with smtplib.SMTP(smtp_server, smtp_port) as server: 
            server.starttls() 
            server.login(sender_email, sender_password) 
            server.sendmail(sender_email, recipient_email, msg.as_string())




