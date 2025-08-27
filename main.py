import os
import pandas as pd

caminho = r"C:\Users\Alexandre-PC\projetos\juntando-planilha\data"

#Lista todos os arquivos na pasta
arquivos = [os.path.join(caminho, arquivo) for arquivo in os.listdir(caminho)]

tabela = pd.DataFrame()
for arquivo in arquivos:
    df = pd.read_excel(arquivo, index_col="NFS_EMISSAO")
    tabela = pd.concat([tabela, df])

tabela.to_excel("dados_das_planilhas.xlsx")
