import pandas as pd
import configparser
from sqlalchemy import create_engine
from sqlalchemy.engine.url import URL
from getpass import getpass

config = configparser.ConfigParser()
config.read('config.ini')

host = config['Conexao']['host']
database = config['Conexao']['database']
#username = config['Conexao']['username']
#password = config['Conexao']['password']

username = "PRONIM"
password = "PRO98NIM"

SERVER = host
DATABASE = database
USERNAME = username
PASSWORD = password

connectionString = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
conn_url = URL.create("mssql+pyodbc", query={"odbc_connect": connectionString})
engine = create_engine(conn_url)

print(f"Conexão à base {DATABASE} realizada com sucesso!")

sql_query = pd.read_sql_query(
    """ 
    SELECT
        '06/06/2024' as [Data Movimento],	
        ITEM.nrPlaca as Placa,
        ITEM.dsItem as [Descrição Completa],
        ITEM.dsReduzida as [Descrição Reduzida],
        ITEM.dsMarca AS MARCA,
        ITEM.dsModelo AS MODELO,
        ITEM.cdLocalizacao as [Cód Localização],
        ITEM.cdClassificacao as [Cód. Classificação],
        ITEM.cdSituacao as Situação,
        ITEM.cdEstadoConser as EstadoConservação,
        Convert(varchar(10), dtAquisicao, 105) as [Data do Ingresso],
        ITEM.cdTpIngresso as [Tipo de Ingresso],
        ITEM.cdFornecedor as [Cód. Fornecedor],
        ITEM.cdConvenio as Convenio,
        ITEM.vlAtual as [Valor Atual],
        ITEM.InContabil as Contábil,
        ITEM.InDepreciavel as Depreciável,
        ITEM.CdMetodoDepreciacao as [Método de Depreciação],
        ITEM.VidaUtil as [Vida Útil],
        ITEM.vlResidual as [Valor Residual],
        '06/06/2024' as [Data Movimento]
    FROM ITEM
    WHERE ITEM.stitem = 'N'
    """,
    engine
)

df = pd.DataFrame(sql_query)

df.to_excel(r"C:\Temp\Itens_Levantamento.xlsx", index=False)

print("Geração do arquivo 'Itens_Levantamento.xlsx' realizada com Sucesso!")
