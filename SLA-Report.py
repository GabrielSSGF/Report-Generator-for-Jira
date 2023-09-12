import pandas as pd
import datetime
import requests
import json
from requests.auth import HTTPBasicAuth
from datetime import date, datetime

def main():
    global dadosJson
    dadosJson = getDataFromJsonFile()
    
    dataFrame = jsonToDataFrame(getApiJsonSLA())
    
    dataFrameSLAMensal = createSLADataFrame(dataFrame, 'relatórioSLA_Semanal')
    dataFrameSLASemanal = createSLADataFrame(dataFrame, 'relatórioSLA_Mensal')

    dataFramesFinais = [dataFrameSLASemanal, dataFrameSLAMensal]

    exportacaoXLSX(dataFramesFinais, ['relatórioSLA_Semanal', 'relatórioSLA_Mensal'])
    
def getDataFromJsonFile():
    with open('configData.json', 'r') as arquivoJson:
            dadosJson = json.load(arquivoJson)
    return dadosJson

def getApiJsonSLA():
    global AUTH
    
    AUTH = HTTPBasicAuth(dadosJson["Email"], dadosJson["APIToken"])
    url = f'https://{dadosJson["CompanyDomain"]}.atlassian.net/rest/api/3/search'
    jql = dadosJson["JQLSLA"]

    headers = {
        "Accept": "application/json"
    }

    params = {
        "jql": jql,
        "maxResults": 100
    }

    response = requests.get(url, headers=headers, auth=AUTH, params=params)
    
    return response
    
def jsonToDataFrame(response):
    listaDados = []
    
    data = response.json()
    issues = data["issues"]

    for issue in issues:
        key = issue["key"]
        created = issue["fields"]["created"]
        timeFirstResponse, timeResolution = getSLAData(key)
        resolutionDate = issue["fields"]["resolutiondate"]

        row = [key, created, timeFirstResponse, timeResolution, resolutionDate]
        listaDados.append(row)
                
    dataFrame = pd.DataFrame(listaDados, columns=["Key", "Created", "Time to resolution", "Time to first response", "Resolved"])        
    
    return dataFrame

def getSLAData(key):
    slaInfo = []
    requestSLAInfo = requestIssueInfo(key)
    dadosSLA = requestSLAInfo['values']
    for item in dadosSLA:
        timeToSomething = getCyclesData(item, item['name'])
        slaInfo.append(timeToSomething)
    return slaInfo

def requestIssueInfo(key):
    url = f'https://{dadosJson["CompanyDomain"]}.atlassian.net/rest/servicedeskapi/request/{key}/sla'
    headers = {
        "Accept": "application/json"
    }
    response = requests.get(url, headers=headers, auth=AUTH)
    return response.json()

def getCyclesData(item, timeType):
    if item['name'] == timeType:
        if item['completedCycles'] and len(item['completedCycles']) > 0: timeToSomething = item['completedCycles'][0]['breached']
        elif 'ongoingCycle' in item: timeToSomething = item['ongoingCycle']['breached']       
    return timeToSomething 

def createSLADataFrame(dataFrame, nomeDataFrame):
    tempo = dataFrame[['Time to first response', 'Time to resolution']].copy()

    dataFrame['Created'] = pd.to_datetime(dataFrame['Created'], format='%Y-%m-%dT%H:%M:%S.%f%z')
    dataFrame['Resolved'] = pd.to_datetime(dataFrame['Resolved'], format='%Y-%m-%dT%H:%M:%S.%f%z')

    nomeDataFrame = nomeDataFrame.split('_')

    if nomeDataFrame[-1] == "Mensal":
        periodo = "Month"
        dataFrame_nova_tabela = tempoRelatorio(periodo, dataFrame, tempo)
        return dataFrame_nova_tabela 

    periodo = 'Week'
    dataFrameNovaTabela = tempoRelatorio(periodo, dataFrame, tempo)
    
    return dataFrameNovaTabela

def tempoRelatorio(periodo, dataFrame, tempo):
    if periodo == "Week":
        dataFrame['Week'] = dataFrame['Resolved'].dt.strftime('%W')
        qntdPorSemana = dataFrame['Week'].value_counts().sort_index()
        periodosSlas = calculoDePorcentagens(tempo, dataFrame, qntdPorSemana, periodo)
        dataFrameNovaTabela = criacaoTabelaSemanal(periodosSlas, qntdPorSemana)
    else:
        dataFrame['Month'] = dataFrame['Resolved'].dt.month
        qntdPorMes = dataFrame['Month'].value_counts()
        
        periodosSlas = calculoDePorcentagens(tempo, dataFrame, qntdPorMes, periodo)
        dataFrameNovaTabela = criacaoTabelaMensal(periodosSlas, qntdPorMes)
                
    return dataFrameNovaTabela
    
def calculoDePorcentagens(tempo, dataFrame, qntd, periodo):
    slaFirstResponseSemQuebra = tempo[tempo['Time to first response'] == False]
    slaResolutionSemQuebra = tempo[tempo['Time to resolution'] == False]

    slaFirstResponse = (slaFirstResponseSemQuebra.groupby(dataFrame[periodo]).size() / qntd * 100).round().astype(int)
    slaResolution = (slaResolutionSemQuebra.groupby(dataFrame[periodo]).size() / qntd * 100).round().astype(int)
    
    return slaFirstResponse, slaResolution

def criacaoTabelaSemanal(slas, qntd):
    slaFirstResponse = slas[0]
    slaResolution = slas[1]
    
    semanas = slaFirstResponse.index.to_series().astype(str).reset_index(drop=True)
    
    startDates = []

    for semana in semanas:
        week_number = int(semana.split()[0])
        year = date.today().isocalendar()[0]
        startDate = datetime.strptime(f'{year}-W{week_number}-1', "%Y-W%W-%w").date()
        startDates.append(startDate.strftime("%d-%m-%Y"))
    
    allDates = pd.date_range(start=startDates[0], end=pd.Timestamp.today(), freq='7D').strftime("%d-%m-%Y").tolist()
    
    data = {}
    for dateString in allDates:
        if dateString in startDates:
            index = startDates.index(dateString)
            slaFirstResponseValor = slaFirstResponse[index]
            slaResolutionValor = slaResolution[index]
            ticketsResolvidosValor = qntd[index] if pd.notnull(qntd[index]) else '-'
            data[dateString] = [slaFirstResponseValor, slaResolutionValor, ticketsResolvidosValor]
        else:
            data[dateString] = ['-', '-', '-']
            
    dataFrame = pd.DataFrame(data, index=['SLA First Response', 'SLA Resolution', 'Tickets Resolvidos'])
            
    return dataFrame

def criacaoTabelaMensal(slas, qntd):
    slaFirstResponse = slas[0]
    slaResolution = slas[1]
    
    year = date.today().isocalendar()[0]
    
    meses = [f'Janeiro-{year}', f'Fevereiro-{year}', f'Março-{year}', f'Abril-{year}', f'Maio-{year}', f'Junho-{year}', 
            f'Julho-{year}', f'Agosto-{year}', f'Setembro-{year}', f'Outubro-{year}', f'Novembro-{year}', f'Dezembro-{year}']

    data = {}

    for i, mes in enumerate(meses, start=1):
        if i in slaFirstResponse.index and pd.notnull(slaFirstResponse[i]):
            data[mes] = [slaFirstResponse[i], slaResolution[i], qntd[i]]

    dataFrame = pd.DataFrame(data, index=['SLA First Response', 'SLA Resolution', 'Tickets Resolvidos'])
            
    return dataFrame

def exportacaoXLSX(dataFrames, planilhas):
    dataAtual = datetime.now()
    diaAtual = dataAtual.strftime('%d')
    mesAtual = dataAtual.strftime("%m")
    anoAtual = dataAtual.year
    
    caminhoRelatorio = f'{dadosJson["FilePathSLA"]}_{diaAtual}-{mesAtual}-{anoAtual}.xlsx'
    
    writer = pd.ExcelWriter(caminhoRelatorio, engine='openpyxl')
    for i, dataFrame in enumerate(dataFrames):
        dataFrame.to_excel(writer, sheet_name=planilhas[i])
        workbook = writer.book
        worksheet = workbook[planilhas[i]]
        padraoDeColunas(worksheet)

    writer._save()        

def padraoDeColunas(worksheet):
    colunas = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'R', 
            'S', 'T', 'V', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI']
    
    worksheet.column_dimensions['A'].width = 22
    worksheet.column_dimensions['B'].width = 22

    for i in range(len(colunas)):
        worksheet.column_dimensions[colunas[i]].width = 20
            
if __name__ == '__main__':
    main()
