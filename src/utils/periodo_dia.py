from datetime import datetime

def class_periodo_dia(valor):
    valor = str(valor)
    if len(valor) == 8 and ":" in valor:
        hr = datetime.strptime(valor,'%H:%M:%S')

        if datetime.strptime('18:00:00', '%H:%M:%S').time() < hr.time():
            return 'NOITE'
        elif datetime.strptime('12:00:00', '%H:%M:%S').time() < hr.time():
            return 'TARDE'
        elif datetime.strptime('06:00:00', '%H:%M:%S').time() < hr.time():
            return 'MANHÃ'
        return 'MADRUGADA'
    return 'S/R'