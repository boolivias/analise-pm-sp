import os
import pandas as pd
import numpy as np

def maior_ocorrencia(x):
    df = x.value_counts(dropna=False)
    df_max = df[df == df.max()].index.array
    return ', '.join('NA' if value is np.nan else value for value in df_max)

# Diretório base contendo as pastas dos anos
diretorio_base = './files/output'

# Lista de anos (pastas) a serem percorridos
anos = [str(ano) for ano in range(2013, 2023)]

# Para cada ano (pasta) no diretório base
for ano in anos:
    # Dicionário para armazenar as analises de cada cluster
    analise_cluster = {}

    # Caminho completo para o diretório do ano
    diretorio_ano = f"{diretorio_base}/ {ano}"
    
    # Caminho completo para o arquivo .xlsx
    arquivo_xlsx = f"{diretorio_ano}/kohonen_info.xlsx"

    abas = pd.ExcelFile(arquivo_xlsx).sheet_names[1:] # Obter todas as abas, exceto a primeira (aba de dados)
    aba_dados = pd.read_excel(arquivo_xlsx, sheet_name="Dados")
    aba_dados = aba_dados[['Neuronio', 'COORPORAÇÃO', 'SITUAÇÃO', 'DESC_TIPOLOCAL', 'COR_PELE', 'PROFISSAO', 'REGIAO', 'FAIXA_ETARIA', 'PERIODO_DIA']]
    dados_agrupados = aba_dados.groupby(['Neuronio']).agg(maior_ocorrencia).reset_index()
    
    writer = pd.ExcelWriter(f"{diretorio_ano}/kohonen_analise.xlsx")
    for aba in abas:
        sheet = pd.read_excel(arquivo_xlsx, sheet_name=aba)
        cluster_analise = dados_agrupados[dados_agrupados['Neuronio'].isin(sheet['Neuronio'].values)]
        cluster_analise.to_excel(writer, sheet_name=aba)
    writer.close()
