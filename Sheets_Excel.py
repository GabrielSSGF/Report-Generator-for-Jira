import pandas as pd
import datetime
from datetime import date, datetime

"""
    Funções usadas em ambas
"""

def exportacao_xlsx(data_frames, planilhas, caminho_relatorio, tipo_relatorio):
    writer = pd.ExcelWriter(caminho_relatorio, engine='openpyxl')

    for i, df in enumerate(data_frames):
        df.to_excel(writer, sheet_name=planilhas[i])

        workbook = writer.book
        worksheet = workbook[planilhas[i]]
        if tipo_relatorio == 0:
            padrao_de_colunas_SLA(worksheet)
            continue
        padrao_de_colunas_Status(worksheet)

    writer._save()


""" 
    SLA
"""

def padrao_de_colunas_SLA(worksheet):
    colunas = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'R', 
               'S', 'T', 'V', 'Z', 'AA', 'AB', 'AC', 'AD']
    
    worksheet.column_dimensions['A'].width = 22
    worksheet.column_dimensions['B'].width = 22

    for i in range(len(colunas)):
        worksheet.column_dimensions[colunas[i]].width = 20

def criacao_tabela_mensal(slas, qntd):
    sla_firstResponse = slas[0]
    sla_resolution = slas[1]
    year = date.today().isocalendar()[0]
    
    meses = [f'Janeiro-{year}', f'Fevereiro-{year}', f'Março-{year}', f'Abril-{year}', f'Maio-{year}', f'Junho-{year}', 
             f'Julho-{year}', f'Agosto-{year}', f'Setembro-{year}', f'Outubro-{year}', f'Novembro-{year}', f'Dezembro-{year}']

    data = {}

    for i, mes in enumerate(meses, start=1):
        if i in sla_firstResponse.index and pd.notnull(sla_firstResponse[i]):
            data[mes] = [sla_firstResponse[i], sla_resolution[i], qntd[i]]

    df = pd.DataFrame(data, index=['SLA First Response', 'SLA Resolution', 'Tickets Resolvidos'])
    return df

def criacao_tabela_semanal(slas, qntd):
    sla_firstResponse = slas[0]
    sla_resolution = slas[1]
    
    semanas = sla_firstResponse.index.to_series().astype(str).reset_index(drop=True)
    
    start_dates = []

    for semana in semanas:
        week_number = int(semana.split()[0])
        year = date.today().isocalendar()[0]
        start_date = datetime.strptime(f'{year}-W{week_number}-1', "%Y-W%W-%w").date()
        start_dates.append(start_date.strftime("%d-%m-%Y"))
    
    data = {}
    for i, semana in enumerate(start_dates):
        if pd.notnull(sla_firstResponse[i]):
            data[semana] = [sla_firstResponse[i], sla_resolution[i], qntd[i]]

    df = pd.DataFrame(data, index=['SLA First Response', 'SLA Resolution', 'Tickets Resolvidos'])
    return df


""" 
    Status
"""

def padrao_de_colunas_Status(worksheet):
    colunas = ['D', 'H', 'M', 'N', 
               'S', 'T', 'V', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']
    
    colunas_large = ['C', 'E', 'G', 'I', 'J', 'O', 'P', 'R', 'U', 'W', 
                     'F', 'Q', 'K', 'L', 'R', 'X', 'Y']
    
    worksheet.column_dimensions['A'].width = 13
    worksheet.column_dimensions['B'].width = 62

    for i in range(len(colunas_large)):
        worksheet.column_dimensions[colunas_large[i]].width = 30

    for i in range(len(colunas)):
        worksheet.column_dimensions[colunas[i]].width = 20

def estilo_negrito(row, df):
    if row.name == df.index[-1]:
        return ['font-weight: bold'] * len(row)
    return [''] * len(row)
