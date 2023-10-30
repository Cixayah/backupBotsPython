import pandas as pd
from sqlalchemy import create_engine

# Cria a string de conexão com o banco de dados
db_connection_str = "mysql+mysqldb://thiago:informatica@andreserver/pjedados"

# Cria a conexão com o banco de dados
db_connection = create_engine(db_connection_str)

# Lê os dados da planilha do Excel
df = pd.read_excel(
    r"C:\Users\Gabriel V Costa\Desktop\excel_sql\adicionar ao Banco\relatoriobot.xlsx"
)

# Converte o formato da data nas colunas dtAutuacao e dtConsulta
df["dtAutuacao"] = pd.to_datetime(
    df["dtConsulta"], format="%d-%m-%Y %H:%M:%S"
).dt.strftime("%Y-%m-%d %H:%M:%S")
df["ultimoMovimento"] = pd.to_datetime(
    df["dtConsulta"], format="%d-%m-%Y %H:%M:%S"
).dt.strftime("%Y-%m-%d %H:%M:%S")

df["dtConsulta"] = pd.to_datetime(
    df["dtConsulta"], format="%d-%m-%Y %H:%M:%S"
).dt.strftime("%Y-%m-%d %H:%M:%S")

# Insere os dados no banco de dados
df.to_sql(name="eprocconsulta", con=db_connection, if_exists="append", index=False)
