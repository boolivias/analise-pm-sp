import pandas as pd
import numpy as np

def maior_ocorrencia(x):
    df = x.value_counts(dropna=False)
    df_max = df[df == df.max()].index.array
    return ', '.join('NA' if value is np.nan else value for value in df_max)

def formata_maior_media(serie, estatistica_df):
    if serie.name not in estatistica_df.columns:
        return ['' for v in serie]
    array_verificado = serie >= estatistica_df[serie.name][0]
    return ['color: red' if v else '' for v in array_verificado]

# Diretório base contendo as pastas dos anos
diretorio_base = '../out'

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
    aba_dados = aba_dados[['Neuronio', 'COR_PELE', 'FAIXA_ETARIA', 'PERIODO_DIA']]

    # dados_agrupados = aba_dados.groupby(['Neuronio']).agg(maior_ocorrencia).reset_index()
    # writer = pd.ExcelWriter(f"{diretorio_ano}/kohonen_analise.xlsx")
    # for aba in abas:
    #     sheet = pd.read_excel(arquivo_xlsx, sheet_name=aba)
    #     cluster_analise = dados_agrupados[dados_agrupados['Neuronio'].isin(sheet['neuronio'].values)]
    #     cluster_analise.to_excel(writer, sheet_name=aba)
    # writer.close()

    writer = pd.ExcelWriter(f"{diretorio_ano}/kohonen_analise.xlsx")
    for aba in abas:
        df_clusters = pd.read_excel(arquivo_xlsx, sheet_name=aba)
        estatistica_df = pd.DataFrame()
        for idx, coluna in enumerate(df_clusters.columns):
            if coluna in ['neuronio', 'cluster']: continue
            dados = df_clusters[coluna]
            media = dados.mean()
            mediana = dados.median()
            variancia = dados.var()
            desv_padrao = dados.std()
            estatistica_df[coluna] = [media, mediana, variancia, desv_padrao]
        df_clusters.insert(0, 'qt_registros', df_clusters['neuronio'].apply(lambda x: len(aba_dados[aba_dados['Neuronio'] == x])))
        df_clusters.style.apply(formata_maior_media, estatistica_df=estatistica_df).to_excel(writer, sheet_name=aba, index=False)

        estatistica_df.insert(0, 'Estatistica', ['Média', 'Mediana', 'Variância', 'Desvio Padrão'])
        estatistica_df.to_excel(writer, sheet_name=aba, startrow=len(df_clusters)+2, index=False, header=False)
    writer.close()

    writer = pd.ExcelWriter(f"{diretorio_ano}/kohonen_maiores_vetores.xlsx")

    for aba in abas:
        df_clusters = pd.read_excel(arquivo_xlsx, sheet_name=aba)
        maiores = df_clusters.drop(['neuronio', 'cluster'], axis=1).apply(
            lambda row: row.nlargest(3).index.tolist(),
            axis = 1
        )
        df_maiores = pd.DataFrame(maiores.tolist(), columns=['maior_valor_1', 'maior_valor_2', 'maior_valor_3'])
        # df['neuronio'] = df_clusters['neuronio']
        df = pd.concat([df_clusters['neuronio'], df_maiores], axis=1)
        # df.columns = ['neuronio', 'maior_valor_1', 'maior_valor_2', 'maior_valor_3']
        df.to_excel(writer, sheet_name=aba, index=False)

        worksheet = writer.sheets[aba]
        for idx, col in enumerate(df.columns):
            series = df[col]
            tam_max = max((series.astype(str).apply(len).max(), len(str(col)))) + 2
            worksheet.set_column(idx, idx, tam_max)
        for row in range(1, len(df) + 1):
            worksheet.set_row(row, None, writer.book.add_format({'top': 1, 'bottom': 1}))  # 1 para borda fina
    writer.close()

    writer = pd.ExcelWriter(f"{diretorio_ano}/kohonen_neuronios.xlsx")
    for neuronio in range(1, 36):
        dados_neuronio = aba_dados[aba_dados['Neuronio'] == neuronio]
        dados_neuronio.to_excel(writer, sheet_name=f"Neuronio {neuronio} ({len(dados_neuronio)})")
    writer.close()