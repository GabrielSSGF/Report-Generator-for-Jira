import pandas as pd
import Calculos
import pytz
import Sheets_Excel
from datetime import datetime
from dateutil.parser import parse


""" 
    SLA
"""

def tempo_relatorio(periodo, df, tempo):
    if periodo == "Week":
        df['Week'] = df['Resolved'].dt.strftime('%W')
        qntd_por_semana = df['Week'].value_counts().sort_index()
        periodos_slas = Calculos.calculo_de_porcentagens(tempo, df, qntd_por_semana, periodo)
        df_nova_tabela = Sheets_Excel.criacao_tabela_semanal(periodos_slas, qntd_por_semana)
    else:
        df['Month'] = df['Resolved'].dt.month
        qntd_por_mes = df['Month'].value_counts()
        
        periodos_slas = Calculos.calculo_de_porcentagens(tempo, df, qntd_por_mes, periodo)
        df_nova_tabela = Sheets_Excel.criacao_tabela_mensal(periodos_slas, qntd_por_mes)
        
    return df_nova_tabela

def relatorio_SLA(df, nome_df):

    tempo = df[['Time to first response', 'Time to resolution']].copy()

    tempo['Time to first response'] = pd.to_timedelta(tempo['Time to first response'])
    tempo['Time to resolution'] = pd.to_timedelta(tempo['Time to resolution'])

    df['Created'] = pd.to_datetime(df['Created'], format='%Y-%m-%dT%H:%M:%S.%f%z')
    df['Resolved'] = pd.to_datetime(df['Resolved'], format='%Y-%m-%dT%H:%M:%S.%f%z')

    nome_df = nome_df.split('_')

    if nome_df[-1] == "Mensal":
        periodo = "Month"
        df_nova_tabela = tempo_relatorio(periodo, df, tempo)
        return df_nova_tabela 

    periodo = 'Week'
    df_nova_tabela = tempo_relatorio(periodo, df, tempo)

    return df_nova_tabela

""" 
    Status
"""

def relatorio_Status(df):
    df['Status Transition.date'] = df['Status Transition.date'].apply(lambda x: x if pd.notnull(x) else '')
    df['next_transition_date'] = df.groupby('Key')['Status Transition.date'].shift(-1)

    current_datetime = datetime.now(pytz.timezone("INSIRA SUA TIMEZONE")) # Insira sua Timezone. Ex: America/Sao_Paulo

    current_date = current_datetime.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + current_datetime.strftime("%z")

    df.loc[df['Status Transition.to'] != 'Concluído', 'next_transition_date'] = df.loc[df['Status Transition.to'] != 'Concluído', 'next_transition_date'].fillna(current_date)

    df['Status Transition.date'] = df['Status Transition.date'].apply(lambda x: parse(x, fuzzy=True) if pd.notnull(x) else x)
    df['next_transition_date'] = df['next_transition_date'].apply(lambda x: parse(x, fuzzy=True) if pd.notnull(x) else x)

    # Remover informações de fuso horário dos objetos datetime
    df['Status Transition.date'] = df['Status Transition.date'].apply(lambda x: x.replace(tzinfo=None) if pd.notnull(x) else x)
    df['next_transition_date'] = df['next_transition_date'].apply(lambda x: x.replace(tzinfo=None) if pd.notnull(x) else x)

    df['Time Interval'] = df.apply(lambda row: Calculos.subtrair_datas(row['Status Transition.date'], row['next_transition_date']), axis=1)

    valores_totais = ['Total', 'Total Group 1', 'Total Group 2', 'Total Group 3']
    total = Calculos.calculo_de_total_e_slas(df)

    df_pivot = df.pivot_table(index=['Key', 'Time to first response', 'Time to resolution'],
                          columns='Status Transition.to',
                          values='Time Interval',
                          aggfunc='sum')

    for i in range(len(valores_totais)):
        df_pivot[valores_totais[i]] = total[i]

    colunas_numericas = df_pivot.select_dtypes(include='number')
    medias = colunas_numericas.mean()
    df_pivot = df_pivot.applymap(Calculos.converter_para_horas_minutos)

    df_pivot = df_pivot.reset_index()

    linha_nova = pd.Series(medias, name='Média')
    
    # SLAS
    
    df_pivot['SLA First Response'] = df_pivot['Time to first response'].apply(lambda x: "Quebrado" if pd.notnull(x) and str(x).startswith('-') else "Cumprido")
    df_pivot['SLA Resolution'] = df_pivot['Time to resolution'].apply(lambda x: "Quebrado" if pd.notnull(x) and str(x).startswith('-') else "Cumprido")
    colunas_time = ['Time to resolution', 'Time to first response']
    df_pivot = df_pivot.drop(colunas_time, axis=1)

    # Média

    df_pivot = df_pivot._append(linha_nova)
    df_pivot.loc['Média'] = df_pivot.loc['Média'].apply(Calculos.converter_para_horas_minutos)
    
    df_pivot = df_pivot.set_index('Key').sort_values(by='Key', ascending=True)
    df_pivot = df_pivot.reindex(df_pivot.index[1:].tolist() + [df_pivot.index[0]])
    df_pivot.rename(index={"-": "Média"}, inplace=True)

    df_pivot = df_pivot.style.apply(Sheets_Excel.estilo_negrito, df=df_pivot, axis=1)

    return df_pivot
