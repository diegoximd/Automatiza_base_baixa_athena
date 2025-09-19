import firebirdsql
import pandas as pd

conn = firebirdsql.connect(
    host='servidor2',
    database=r'D:\Dados_interbase\COB_DB_EXECUTIVA_ATHENA_SAUDE.FDB',
    user='consulta',
    password='@BmpAdm35ConsultaSql#',
    charset='UTF8'
)

cur = conn.cursor()
cur.execute("SELECT FIRST 100 NROPERACAO, BANCO FROM OPERACOES WHERE BANCO=2002")
dados = cur.fetchall()
colunas = [desc[0] for desc in cur.description]
df_banco = pd.DataFrame(dados, columns=colunas)

cur.close()
conn.close()

print(df_banco.head())
