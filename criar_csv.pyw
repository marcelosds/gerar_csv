import logging
import sys
import urllib
import webbrowser
import pandas as pd
import configparser
import pyodbc
from tkinter import *
import os
from tkinter import messagebox
from sqlalchemy import create_engine
from PyQt5.QtGui import QIcon


class Application:
    def __init__(self, master=None):
        self.fontePadrao = ("Arial", "10")
        self.primeiroContainer = Frame(master)
        self.primeiroContainer["pady"] = 10
        self.primeiroContainer.pack()

        self.segundoContainer = Frame(master)
        self.segundoContainer["padx"] = 20
        self.segundoContainer.pack()

        self.terceiroContainer = Frame(master)
        self.terceiroContainer["padx"] = 20
        self.terceiroContainer.pack()

        self.quartoContainer = Frame(master)
        self.quartoContainer["pady"] = 20
        self.quartoContainer.pack()

        self.titulo = Label(self.primeiroContainer, text="Autenticar na Base de Dados")
        self.titulo["font"] = ("Arial", "10", "bold")
        self.titulo.pack()

        self.nomeLabel = Label(self.segundoContainer, text="Usuário: ", font=self.fontePadrao)
        self.nomeLabel.pack(side=LEFT)

        self.nome = Entry(self.segundoContainer)
        self.nome["width"] = 30
        self.nome["font"] = self.fontePadrao
        self.nome.pack(side=LEFT)

        self.senhaLabel = Label(self.terceiroContainer, text="Senha: ", font=self.fontePadrao)
        self.senhaLabel.pack(side=LEFT)

        self.senha = Entry(self.terceiroContainer)
        self.senha["width"] = 30
        self.senha["font"] = self.fontePadrao
        self.senha["show"] = "*"
        self.senha.pack(side=LEFT)

        self.autenticar = Button(self.quartoContainer)
        self.autenticar["text"] = "Gerar"
        self.autenticar["font"] = ("Calibri", "8")
        self.autenticar["width"] = 12
        self.autenticar["command"] = self.gerarItens
        self.autenticar.pack()


    #Método Autenticar e Gerar Lista de Itens
    def gerarItens(self):

        # Verificar se o driver está instalado
        drivers = [driver for driver in pyodbc.drivers() if "ODBC Driver 13 for SQL Server" in driver]
        #lista_drivers = [driver for driver in pyodbc.drivers()]
        #print(lista_drivers)

        if drivers:
            try:
                config = configparser.ConfigParser()
                config.read('config.ini')

                host = config['Conexao']['host']
                database = config['Conexao']['database']

                usuario = self.nome.get()
                senha = self.senha.get()

                DRIVER = 'ODBC Driver 13 for SQL Server'
                SERVER = host
                DATABASE = database
                USERNAME = usuario
                PASSWORD = senha

                params = urllib.parse.quote_plus(f"DRIVER={DRIVER};"
                                                 f"SERVER={SERVER};"
                                                 f"DATABASE={DATABASE};"
                                                 f"UID={USERNAME};"
                                                 f"PWD={PASSWORD}")

                engine = create_engine("mssql+pyodbc:///?odbc_connect={}".format(params))


                query = """
                        SELECT
                            '' as [Data Movimento],
                            ITEM.cditem as [Código],	
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
                            '' as [Data Inicio Depreciação]
                        FROM ITEM
                        WHERE ITEM.stitem = 'N'
                        """

                df = pd.read_sql(query, engine)

                diretorio_atual = os.getcwd()
                caminho_arquivo = os.path.join(diretorio_atual, "Itens_Levantamento.xlsx")

                df.to_excel(caminho_arquivo, index=False)

                messagebox.showinfo('Gerar Lista', 'Lista gerada com sucesso!')

            except Exception as e:
                logging.basicConfig(filename='app.log', level=logging.ERROR,
                                    format='%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s')
                logging.error('{}'.format(str(e)))
                messagebox.showwarning('Conexão SQL', 'Ocorreu um erro de conexão, verifique o arquivo de LOG.')

        else:
            messagebox.showwarning('Driver ODBC', 'Driver ODBC 13 for SQL Server não instalado, por favor instale!')
            # URL que você deseja abrir
            url = "https://www.microsoft.com/en-us/download/details.aspx?id=50420"
            # Abrir o navegador padrão com o link
            webbrowser.open(url)


app = Tk()
app.geometry("300x160")
app.eval('tk::PlaceWindow %s center' % app.winfo_pathname(app.winfo_id()))
app.title("Gerar Itens - GOVBR PP")
app.iconbitmap("icon.ico")
app.resizable(0, 0)
Application(app)
app.mainloop()

if getattr(sys, 'frozen', False):
    # Se o código estiver sendo executado em um executável PyInstaller
    app_icon = QIcon('icon.ico')
    app.setWindowIcon(app_icon)