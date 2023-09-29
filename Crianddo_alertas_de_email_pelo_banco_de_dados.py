
import configparser
import datetime
import warnings
import pandas as pd
import smtplib
import pyodbc
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

#Ignorar warning do pyodbc
warnings.filterwarnings('ignore')

# Variavel que le o arquivo de configuracao
config = configparser.ConfigParser()
config.read('infoDB.ini')

# Dados de acesso ao Banco de Dados
BD_USER = str(config['SQLSERVER']['BD_USER'])
BD_PASS = str(config['SQLSERVER']['BD_PASS'])
BD_HOST = str(config['SQLSERVER']['BD_HOST'])
BD_NAME = str(config['SQLSERVER']['BD_NAME'])

# Querys padrao
query_01 = """<<INSERIR QUERY>>"""

query_02 = """<<INSERIR QUERY>>"""


def retornaDataHora():
  return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") 

def conexao():
  conn_str = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + BD_HOST + ';DATABASE=' + BD_NAME + ';UID=' + BD_USER + ';PWD=' + BD_PASS + ';Trusted_Connection=no;'
  # Conexao com o banco de dados
  cnxn = pyodbc.connect(conn_str)
  cursor = cnxn.cursor()
  return cnxn, cursor

def executaQuery(texto, tipo):
  cnxn, cursor = conexao()
  resultado = pd.read_sql(texto, cnxn)
  if tipo == 1:
    grouped = resultado.groupby('id_deal')
    cursor = cnxn.cursor()
    return resultado, grouped
  else:
    return resultado


def enviaEmail():
  smtp_server = "outlook.office365.com"
  smtp_port = 587
  sender_email = "<<EMAIL>>"
  sender_password = "<<SENHA>>"
  nk_deal = 0
  totalAlmoxarifadosNovos = (executaQuery(query_01,0)).iloc[0]["Novo"]

  if totalAlmoxarifadosNovos!=0:
    mensagemTotal = f"<<Mensagem>>"
    list_almoxarifado = pd.DataFrame(executaQuery(query_02,0))
    html = list_almoxarifado.to_html(index=False)
    corpo = f"""<p><<Mensagem>></p>            
            <p><<Mensagem>></p>
            <p>{mensagemTotal},<<Mensagem>> </p>"""
    body = f"""<html>
          <head></head>
          <body>
            {html}
          </body>
        </html>"""
      #Mensagem
    msg = MIMEMultipart()
    msg['From'] = str(Header('<<header>>', 'utf-8'))
    recipient_email = ['<<email>>']
    message =f'{corpo}{body}'
    msg['Subject'] = f"<<Assunto>>"
    msg.attach(MIMEText(message, 'html'))
       
    #Mensagem
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        print("E-mail enviado com sucesso")

def main():
  print(retornaDataHora() + " - Iniciando sistema automacao_almoxarifado. ...")
  # Verifica e envia alerta por e-mail
  enviaEmail()
  print(retornaDataHora() + " - Finalizando sistema automacao_almoxarifado. ...")

if __name__ == "__main__":
    main()
