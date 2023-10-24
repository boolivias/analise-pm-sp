import os
import json
import pandas as pd
from utils.periodo_dia import PeriodoHelper
from utils.faixa_etaria import IdadeHelper

periodo_helper = PeriodoHelper()
idade_helper = IdadeHelper()

excel_writer = pd.ExcelWriter('../data/amostra.xlsx', engine="xlsxwriter")
with pd.ExcelFile("../data/MDIP_2013_2023.xlsx") as excel_reader:
    for sheet in excel_reader.sheet_names:
        try:
            excel_df = pd.read_excel(excel_reader, sheet_name=sheet)
        except ValueError as e:
            print("Script executado com sucesso!")
        except Exception as e:
            print("Error:", e)
        else:
            errors = []
            output = []
            for index, linha in excel_df.iterrows():
                if linha["MUNICIPIO_CIRCUNSCRICAO"] != "S.PAULO":
                    continue

                lat = float(linha["LATITUDE"])
                long = float(linha["LONGITUDE"])

                if lat == 0 or pd.isna(lat) or long == 0 or pd.isna(long): continue

                dt = linha.to_dict()
                dt['FAIXA_ETARIA'] = idade_helper.class_idade(dt['IDADE_PESSOA'])
                dt['PERIODO_DIA'] = periodo_helper.class_periodo_dia(dt['HORA_FATO'])
                output.append(dt)
            with open("../data/erros_preprocessamento.json", "w") as json_out:
                json_out.write(json.dumps(errors, indent=2))

            index = index + 1
            df_output = pd.DataFrame(output)
            df_output.to_excel(excel_writer, sheet_name=sheet)
    excel_writer.close()
