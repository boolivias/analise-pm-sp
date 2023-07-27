import os
import pandas as pd

# Diretório base contendo as pastas dos anos
diretorio_base = './files/output'

# Lista de anos (pastas) a serem percorridos
anos = [str(ano) for ano in range(2013, 2023)]
print(anos)

# Para cada ano (pasta) no diretório base
for ano in anos:
    # Caminho completo para o diretório do ano
    diretorio_ano = f"{diretorio_base}/ {ano}"
    
    # Caminho completo para o arquivo .xlsx
    arquivo_xlsx = f"{diretorio_ano}/kohonen_info.xlsx"
    print(arquivo_xlsx)
    
    # Carregar o arquivo .xlsx em um objeto ExcelFile
    xls = pd.ExcelFile(arquivo_xlsx)
    
    # Obter todas as abas, exceto a primeira
    abas = xls.sheet_names[1:]
    
    # Para cada aba no arquivo .xlsx
    for aba in abas:
        # Carregar a aba em um DataFrame
        df = pd.read_excel(arquivo_xlsx, sheet_name=aba)
        
        # Criar um novo DataFrame vazio para armazenar as tabelas
        df_tabelas = pd.DataFrame()
        
        # Criar uma lista para armazenar as tabelas
        tabelas = []
        
        # Para cada linha do DataFrame
        for _, row in df.iterrows():
            # Obtenha as 5 colunas com os maiores valores
            top_columns = row.nlargest(6)
            
            # Crie uma tabela com nome da coluna e valor
            table = pd.DataFrame({
                'Coluna': top_columns.index,
                'Valor': top_columns.values
            })
            
            # Adicione a tabela à lista de tabelas
            tabelas.append(table)
        
        # Gerar o nome do arquivo .txt com base na aba
        nome_arquivo_txt = f'{aba}.txt'
        
        # Caminho completo para o arquivo .txt
        caminho_arquivo_txt = f'{diretorio_ano}/{nome_arquivo_txt}'
        
        # Escrever o DataFrame de tabelas no arquivo .txt
        df_tabelas.to_csv(caminho_arquivo_txt, sep='\t', index=False)
        with open(caminho_arquivo_txt, 'w') as file:
                for tabela in tabelas:
                    tabela.to_string(file, index=False)
                    file.write('\n\n')
