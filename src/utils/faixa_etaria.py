import pandas as pd

def class_idade(valor):
    try:
        idade = float(valor)
        if not pd.isna(idade):
            if idade > 40.0:
                return '+40'
            elif idade >= 30:
                return '30 a 39'
            elif idade >= 20.0:
                return '20 a 29'
            elif idade >= 10.0:
                return '10 a 19'
            return '-10'
        return 'S/R'
    except Exception as e:
        return 'S/R'