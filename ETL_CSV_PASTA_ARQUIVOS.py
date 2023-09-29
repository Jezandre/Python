import pandas as pd
import pyodbc
import configparser
import datetime
import os

# Variavel que le o arquivo de configuracao
config = configparser.ConfigParser()
config.read('infoDB.ini')

# Dados de acesso ao Banco de Dados local
BD_USER = str(config['DEFAULT']['BD_USER'])
BD_PASS = str(config['DEFAULT']['BD_PASS'])
BD_HOST = str(config['DEFAULT']['BD_HOST'])
BD_BD   = str(config['DEFAULT']['BD_BD'])

def retornaDataHora():
  return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S") 

#Tratamento do dataframe
def load_data_from_csv(file_path):
    mes_ano = pd.read_csv(file_path, nrows=3, header=None)

    return df

#Criar tabela baseada no dataframe
def create_table(cursor, table_name, column_names, column_size=255):
    column_definitions = [f"[{column}] VARCHAR({column_size})" for column in column_names]
    query = f"CREATE TABLE {table_name} ({','.join(column_definitions)}, data_carga VARCHAR(255))"
    cursor.execute(query)
    query_python_user = f'GRANT SELECT, INSERT, UPDATE, DELETE, REFERENCES, ALTER ON {table_name} TO USUARIO'
    cursor.execute(query_python_user)
    print("Tabela Criada com Sucesso")


#Inserir dados no SQL server
def insert_data(cursor, df, table_name):
    placeholders = ','.join(['?'] * len(df.columns))
    sql = f"INSERT INTO {table_name} VALUES ({placeholders}, CONVERT(VARCHAR(19), GETDATE(), 120))"
    values = [tuple(x) for x in df.values]
    cursor.executemany(sql, values)
    print('Valores inseridos com sucesso')


#Analisar todos os arquivos
def pegar_todos_arquivos_csv(pasta, cursor, nome_tabela):
    for nome_arquivo in os.listdir(pasta):
        caminho_completo = os.path.join(pasta, nome_arquivo)        
        if os.path.isfile(caminho_completo):
            df = load_data_from_csv(caminho_completo)
            insert_data(cursor, df, nome_tabela)

# Inserir dados do Csv no banco
def inserirDados():
    # Especificar o caminho do arquivo Excel e o nome da planilha
    pasta = '<<INSERIR CAMINHO DA PASTA>>'
    caminho_arquivo = '<<INSERIR CAMINHO DO ARQUIVO>>'

    # Carregar os dados do Excel para o DataFrame
    df = load_data_from_csv(caminho_arquivo)

    # Criar a string de conexão
    conn_str = 'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + BD_HOST + ';DATABASE=' + BD_BD + ';UID=' + BD_USER + ';PWD=' + BD_PASS + ';Trusted_Connection=no;'

    # Conexão com o banco de dados
    cnxn = pyodbc.connect(conn_str)

    # Exemplo: Executar uma consulta SQL
    cursor = cnxn.cursor()

    # Especificar o nome da tabela no SQL Server
    nome_tabela = '<<NOME DA TABELA>>'

    # Truncar tabela antes da inserção
    sql = f"TRUNCATE TABLE {nome_tabela}"
    cursor.execute(sql)

    # Inserir os dados na tabela
    pegar_todos_arquivos_csv(pasta, cursor, nome_tabela)
    # insert_data(cursor, df, nome_tabela)

    # Confirmar as alterações e encerrar a conexão
    cursor.commit()
    cursor.close()
    
    print("Trabalho concluido")

def main():
  print(retornaDataHora() + " - Iniciando sistema python_csv_absenteismo ...")
  # Verifica planilha e insere dados no SQL Server
  inserirDados()
  print(retornaDataHora() + " - Finalizando sistema python_csv_absenteismo ...")

if __name__ == "__main__":
    main()