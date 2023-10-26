import pandas as pd
import numpy as np
import re
import geopandas as gpd
from shapely.geometry import Point
import matplotlib.pyplot as plt


def plota_mapa(df_dados, arquivo_clusters, abas_clusters, caminho_salva):
    gd_file = gpd.read_file('C:/Users/jean.mendoza/Desktop/IC/projeto/kohonen-map/src/utils/file/SIRGAS_GPKG_distrito.gpkg')
    gd = gd_file.to_crs('epsg:4326')

    _, ax = plt.subplots()
    gd.plot(ax=ax, color='lightgray', edgecolor='black', linewidth=0.2)
    for aba in abas_clusters:
        df_cluster = pd.read_excel(arquivo_clusters, sheet_name=aba)
        neuronios = df_cluster['neuronio'].tolist()
        dados_cluster = df_dados[df_dados['Neuronio'].isin(neuronios)]
        qtd_dados = len(dados_cluster)
        if(qtd_dados == 0): continue

        x = np.array(dados_cluster['LONGITUDE'])
        y = np.array(dados_cluster['LATITUDE'])
        pontos = [Point(x[i], y[i]) for i in range(len(x))]
        df_pontos = gpd.GeoDataFrame(geometry=pontos, crs=gd.crs)
        cor_cluster = re.search(r"Cluster \d+ (\w+)", aba).group(1)
        df_pontos.plot(ax=ax, color=cor_cluster, markersize=5, label=f'{cor_cluster} ({qtd_dados})')

    ax.legend(loc='upper left', bbox_to_anchor=(1, 1))
    ax.set_title('Cidade de São Paulo')
    plt.savefig(caminho_salva, dpi=300)

# diretorio_ano = f"{diretorio_base}/ {ano}"
# arquivo_xlsx = f"{diretorio_ano}/kohonen_info.xlsx"
# abas = pd.ExcelFile(arquivo_xlsx).sheet_names[1:] # Obter todas as abas, exceto a primeira (aba de dados)
# aba_dados = pd.read_excel(arquivo_xlsx, sheet_name="Dados")
# aba_dados = aba_dados[['Neuronio', 'LATITUDE', 'LONGITUDE']]
