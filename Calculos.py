import pandas as pd
import datetime as dt
from datetime import timedelta

"""
    SLA
"""

def calculo_de_porcentagens(tempo, df, qntd, periodo):
    sla_firstResponse = (tempo[tempo['Time to first response'] > pd.Timedelta(0)].groupby(df[periodo]).size() / qntd * 100).map(lambda x: round(x) if pd.notnull(x) else "None")
    sla_resolution = (tempo[tempo['Time to resolution'] > pd.Timedelta(0)].groupby(df[periodo]).size() / qntd * 100).map(lambda x: round(x) if pd.notnull(x) else "None")
    
    return sla_firstResponse, sla_resolution

"""
    Status
"""

def converter_para_horas_minutos(valor):
    if pd.isnull(valor):
        return "-"

    hora = int(valor / 3600)
    minutos = int(valor % 3600) // 60

    return f'{hora:02d}h{minutos:02d}'

def total_por_status(status, data_frame):
    total = data_frame[data_frame['Status Transition.to'].isin(status)].pivot_table(
        index=['Key'],
        columns='Status Transition.to',
        values='Time Interval',
        aggfunc='sum'
    ).sum(axis=1)
    return total

def calculo_de_total_e_slas(df):
    status_grupo1 = ['Status que apontam pra um grupo específico']
    status_grupo2 = ['Status que apontam pra um grupo específico']
    status_grupo3 = ['Status que apontam pra um grupo específico']

    total = df.pivot_table(index=['Key'],
                columns='Status Transition.to',
                values='Time Interval',
                aggfunc='sum').sum(axis=1)
    
    total_grupo1 = total_por_status(status_grupo1, df)
    total_grupo2 = total_por_status(status_grupo2, df)
    total_grupo3 = total_por_status(status_grupo3, df)

    return total, total_grupo1, total_grupo3, total_grupo2

def subtrair_datas(data_inicio, data_fim):
    if pd.isnull(data_inicio) or pd.isnull(data_fim):
        return None

    dias_uteis = 0
    dias_fim_de_semana = 0

    dia_semana = data_inicio.weekday()
    if dia_semana == 5:  # Sábado
        data_inicio += timedelta(days=2)
        data_inicio = data_inicio.replace(hour=9, minute=0)

    elif dia_semana == 6:  # Domingo
        data_inicio += timedelta(days=1)
        data_inicio = data_inicio.replace(hour=9, minute=0)


    if data_inicio.time().hour > 18:
        data_inicio = data_inicio.replace(hour=18, minute=0)

    elif data_inicio.time().hour < 9:
        data_inicio = data_inicio.replace(hour=9, minute=0)


    dia_semana = data_fim.weekday()
    if dia_semana == 5:
        data_fim -= timedelta(days=1)
        data_fim = data_fim.replace(hour=18, minute=0)

    elif dia_semana == 6:
        data_fim -= timedelta(days=2)
        data_fim = data_fim.replace(hour=18, minute=0)


    if data_fim.time().hour > 18:
        data_fim = data_fim.replace(hour=18, minute=0)

    elif data_fim.time().hour < 9:
        data_fim = data_fim.replace(hour=9, minute=0)

    segundos = (data_fim - data_inicio).total_seconds()
    
    if data_inicio.time() > data_fim.time():
        data_fim = data_fim + dt.timedelta(days=1)

    diferenca = data_fim - data_inicio

    for i in range(diferenca.days + 1):
        data = data_inicio + dt.timedelta(days=i)
        dia_semana = data.weekday()

        if dia_semana >= 5:
            dias_fim_de_semana += 1
        else:
            dias_uteis += 1   

    if dias_fim_de_semana > 0:
        segundos_fds = dias_fim_de_semana * 86400
        segundos -= segundos_fds

    if dias_uteis == 2:
        segundos -= 54000
        
    elif dias_uteis > 2:
        segundos_foraServico = (dias_uteis-2) * 54000 + 54000
        segundos -= segundos_foraServico

    return segundos