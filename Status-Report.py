import datetime as dt
from datetime import datetime, timedelta
from dateutil.parser import parse
import json
import openpyxl
import pandas as pd
import pytz
import requests
from requests.auth import HTTPBasicAuth

def main():
    global dadosJson
    dadosJson = getDataFromJsonFile()
    
    dataFrameAberto, dataFrameConcluido, = createDataFrameForAnalysis()

    dataFramePivotResolvidos = pivotDataFrame(dataFrameConcluido)
    dataFramePivotAbertosEResolvidos = pivotDataFrame(dataFrameAberto)
    
    dataFrameToExcel([dataFramePivotResolvidos, dataFramePivotAbertosEResolvidos])

def getDataFromJsonFile():
    with open('configData.json', 'r') as arquivoJson:
        dadosJson = json.load(arquivoJson)
    return dadosJson

def createDataFrameForAnalysis():
    responseAberto, responseConcluido = getJiraAPIResponse()
    return responseJsonToDataframe(responseAberto), responseJsonToDataframe(responseConcluido)

def getJiraAPIResponse():
    url = f'https://{dadosJson["CompanyDomain"]}.atlassian.net/rest/api/3/search'
    
    global AUTH
    
    jqlAberto = dadosJson["JQLAbertoStatus"]
    jqlConcluido = dadosJson["JQLConcluidoStatus"]
    AUTH = HTTPBasicAuth(dadosJson["Email"], dadosJson["APIToken"])

    headers = {
        "Accept": "application/json"
    }

    paramsAberto = {
        "jql": jqlAberto,
        "maxResults": 100
    }

    paramsConcluido = {
        "jql": jqlConcluido,
        "maxResults": 100
    }

    responseAberto = requests.get(url, headers=headers, auth=AUTH, params=paramsAberto)
    responseConcluido = requests.get(url, headers=headers, auth=AUTH, params=paramsConcluido)
    
    return responseAberto, responseConcluido

def responseJsonToDataframe(response):
    data = response.json()
    issues = data['issues']
    
    listaDados = []
    
    for i in issues:
        key = i['key']
        listaDados.extend(getAdditionalIssueInfo(key))
    
    dataFrame = pd.DataFrame(listaDados, columns=["Key", "Status Transition Date", "Status Transition To", "Time to first response", "Time to resolution"])
    dataFrame.set_index("Key", inplace=True)
            
    return dataFrame
        
def getAdditionalIssueInfo(key):
    listaDados = []
    
    requestStatusInfo = requestIssueInfo(key, AUTH, 'status')
    requestSLAInfo = requestIssueInfo(key, AUTH, 'sla') 
    dadosSLA = requestSLAInfo['values']
    dadosTransitions = requestStatusInfo['values']
    slaInfo = []
    
    for item in dadosSLA:
        timeToSomething = getCyclesData(item, item['name'])
        slaInfo.append(timeToSomething)
    
    for item in dadosTransitions:
        statusTransitionTo = item['status']               
        statusTransitionDate = item['statusDate']['jira']
        listaDados.append([key, statusTransitionDate, statusTransitionTo, slaInfo[0], slaInfo[1]])
    
    return listaDados

def requestIssueInfo(key, auth, tipoRequest):
    url = f'https://{dadosJson["CompanyDomain"]}.atlassian.net/rest/servicedeskapi/request/{key}/{tipoRequest}'
    headers = {
        "Accept": "application/json"
    }
    response = requests.get(url, headers=headers, auth=auth)
    return response.json()
        
def getCyclesData(item, timeType):
    if item['name'] == timeType:
        if item['completedCycles'] and len(item['completedCycles']) > 0: timeToSomething = item['completedCycles'][0]['breached']
        elif 'ongoingCycle' in item: timeToSomething = item['ongoingCycle']['breached']       
    return timeToSomething 

def pivotDataFrame(dataFrame):
    createTimeIntervalColumn(dataFrame)
    
    dataFramePivot = dataFrame.pivot_table(index=['Key', 'Time to first response', 'Time to resolution'],
                        columns='Status Transition To',
                        values='Time Interval',
                        aggfunc='sum')
    
    dataFramePivot = getTotalValuesFromGroups(dataFrame, dataFramePivot)
    
    print(dataFramePivot)
    
    colunasNumericas = dataFramePivot.select_dtypes(include='number')
    medias = colunasNumericas.mean()
    dataFramePivot = dataFramePivot.applymap(converterParaHorasMinutos)

    dataFramePivot = dataFramePivot.reset_index()

    novaLinha = pd.Series(medias, name='Média')
    
    dataFramePivot = setSLAColumns(dataFramePivot)
    dataFramePivot = setMeanLine(dataFramePivot, novaLinha)
    dataFramePivot = removeUnnecessaryColumns(dataFramePivot, dadosJson["ColumnsToRemove"])
    
    dataFramePivot = dataFramePivot.style.apply(estiloNegrito, dataFrame=dataFramePivot, axis=1)

    return dataFramePivot

def createTimeIntervalColumn(dataFrame):
    dataFrame['Status Transition Date'] = dataFrame['Status Transition Date'].apply(lambda x: x if pd.notnull(x) else '')
    dataFrame['Previous Transition Date'] = dataFrame.groupby('Key')['Status Transition Date'].shift(+1)

    datetimeAtual = datetime.now(pytz.timezone("America/Sao_Paulo"))
    dataAtual = datetimeAtual.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + datetimeAtual.strftime("%z")

    dataFrame.loc[dataFrame['Status Transition To'] != 'Concluído', 'Previous Transition Date'] = dataFrame.loc[dataFrame['Status Transition To'] != 'Concluído', 'Previous Transition Date'].fillna(dataAtual)

    dataFrame['Status Transition Date'] = pd.to_datetime(dataFrame['Status Transition Date'])
    dataFrame['Previous Transition Date'] = pd.to_datetime(dataFrame['Previous Transition Date'])

    dataFrame['Time Interval'] = dataFrame.apply(lambda row: subtrairDatas(row['Status Transition Date'], row['Previous Transition Date']), axis=1)

def getTotalValuesFromGroups(dataFrame, dataFramePivot):
    valoresTotais = ['Total', 'Total Empresa', 'Total Clientes']
    total = calculoDeTotaleSLAs(dataFrame)
    total = [serie.tolist() for serie in total]
    
    for i in range(len(valoresTotais)):
        dataFramePivot[valoresTotais[i]] = total[i] 
    return dataFramePivot

def subtrairDatas(dataInicio, dataFim):
    if pd.isnull(dataInicio) or pd.isnull(dataFim):
        return None
    
    horarioInicioServico = int(dadosJson["ServiceTimeStart"])
    horarioEncerramentoServico = int(dadosJson["ServiceTimeStop"])
    
    dataInicio = adjustDataInicioFimToServiceTime(dataInicio, horarioInicioServico, horarioEncerramentoServico)
    dataFim = adjustDataInicioFimToServiceTime(dataFim, horarioInicioServico, horarioEncerramentoServico)
   
    segundosTotais = (dataFim - dataInicio).total_seconds()
    
    if dataInicio.time() > dataFim.time():
        dataFim = dataFim + dt.timedelta(days=1)

    totalHorasPorDia = 24
    totalSegundosPorHora = 3600
    diferenca = dataFim - dataInicio
    totalPeriodoServico = horarioEncerramentoServico - horarioInicioServico
    segundosForaPeriodoServico = (totalHorasPorDia - totalPeriodoServico) * totalSegundosPorHora if totalPeriodoServico > 0 else (totalHorasPorDia - totalPeriodoServico*-1) * totalSegundosPorHora
    
    segundosTotais = removeSegundosForaDoServiceTime(segundosTotais, segundosForaPeriodoServico, diferenca, dataInicio)

    return segundosTotais
        
def adjustDataInicioFimToServiceTime(dataTipo, horarioInicioServico, horarioEncerramentoServico):
    sabado = 5
    domingo = 6
    diaSemana = dataTipo.weekday()
    
    if diaSemana == sabado:
        dataTipo += timedelta(days=2)
        dataTipo = dataTipo.replace(hour=horarioInicioServico, minute=0)

    elif diaSemana == domingo:
        dataTipo += timedelta(days=1)
        dataTipo = dataTipo.replace(hour=horarioInicioServico, minute=0)

    if dataTipo.time().hour > horarioEncerramentoServico:
        dataTipo = dataTipo.replace(hour=horarioEncerramentoServico, minute=0)

    elif dataTipo.time().hour < horarioInicioServico:
        dataTipo = dataTipo.replace(hour=horarioInicioServico, minute=0)
    
    return dataTipo

def removeSegundosForaDoServiceTime(segundosTotais, segundosForaPeriodoServico, diferenca, dataInicio):
    diasUteis = 0
    diasFimDeSemana = 0
    
    for i in range(diferenca.days + 1):
        data = dataInicio + dt.timedelta(days=i)
        diaSemana = data.weekday()

        if diaSemana >= 5:
            diasFimDeSemana += 1
        else:
            diasUteis += 1

    if diasFimDeSemana > 0:
        segundosFimDeSemana = diasFimDeSemana * 86400
        segundosTotais -= segundosFimDeSemana

    if diasUteis == 2:
        segundosTotais -= segundosForaPeriodoServico
        
    elif diasUteis > 2:
        segundosForaServico = (diasUteis-2) * segundosForaPeriodoServico + segundosForaPeriodoServico
        segundosTotais -= segundosForaServico
    
    return segundosTotais

def calculoDeTotaleSLAs(dataFrame):
    statusCliente = dadosJson["StatusCustomer"]
    statusEmpresa = dadosJson["StatusCompany"]
            
    total = dataFrame.pivot_table(index='Key',
                columns='Status Transition To',
                values='Time Interval',
                aggfunc='sum').sum(axis=1)
    
    totalCliente = totalPorStatus(statusCliente, dataFrame)
    totalEmpresa = totalPorStatus(statusEmpresa, dataFrame)

    return total, totalEmpresa, totalCliente
        
def totalPorStatus(status, dataFrame):
    total = dataFrame[dataFrame['Status Transition To'].isin(status)].pivot_table(
        index='Key',
        columns='Status Transition To',
        values='Time Interval',
        aggfunc='sum'
    ).sum(axis=1)
    
    return total
        
def setSLAColumns(dataFramePivot):
    dataFramePivot['SLA First Response'] = dataFramePivot['Time to first response'].apply(lambda x: "Quebrado" if pd.notnull(x) and x == True else "Cumprido")
    dataFramePivot['SLA Resolution'] = dataFramePivot['Time to resolution'].apply(lambda x: "Quebrado" if pd.notnull(x) and x == True else "Cumprido")
    colunasTimeTo = ['Time to resolution', 'Time to first response']
    dataFramePivot.drop(colunasTimeTo, axis=1, inplace=True)
    return dataFramePivot

def setMeanLine(dataFramePivot, linhaNova):
    dataFramePivot.loc['Média'] = linhaNova
    dataFramePivot.loc['Média'] = dataFramePivot.loc['Média'].apply(converterParaHorasMinutos)
    dataFramePivot = dataFramePivot.set_index('Key').sort_values(by='Key', ascending=True)
    dataFramePivot = dataFramePivot.reindex(dataFramePivot.index[1:].tolist() + [dataFramePivot.index[0]])
    dataFramePivot.rename(index={"-": "Média"}, inplace=True)
    return dataFramePivot

def removeUnnecessaryColumns(dataFramePivot, columnsToRemove):
    if columnsToRemove:
        for item in columnsToRemove:
            dataFramePivot = checkIfItemIsInDataFrame(item, dataFramePivot)
    return dataFramePivot

def checkIfItemIsInDataFrame(item, dataFrame):
    if item in dataFrame.columns:
        dataFrame.drop(item, axis=1, inplace=True)
    return dataFrame

def converterParaHorasMinutos(valor):
    if pd.isnull(valor):
        return "-"

    hora = int(valor / 3600)
    minutos = int(valor % 3600) // 60

    return f'{hora:02d}h{minutos:02d}'

def estiloNegrito(row, dataFrame):
    if row.name == dataFrame.index[-1]:
        return ['font-weight: bold'] * len(row)
    return [''] * len(row)

def dataFrameToExcel(dataFrames):
    dataAtual = datetime.now()
    diaAtual = dataAtual.strftime('%d')
    mesAtual = dataAtual.strftime("%m")
    anoAtual = dataAtual.year

    caminhoRelatorio = f'{dadosJson["FilePathStatus"]}_{diaAtual}-{mesAtual}-{anoAtual}.xlsx'
    exportacaoXLSX(dataFrames, ['Resolvidos', 'Resolvidos-e-Abertos'], caminhoRelatorio)

def exportacaoXLSX(data_frames, planilhas, caminho_relatorio):
    writer = pd.ExcelWriter(caminho_relatorio, engine='openpyxl')
    for i, dataFrame in enumerate(data_frames):
        dataFrame.to_excel(writer, sheet_name=planilhas[i])
        workbook = writer.book
        worksheet = workbook[planilhas[i]]
        padraoDeColunas(worksheet)
    writer._save()

def padraoDeColunas(worksheet):
    colunas = ['B', 'D', 'H', 'M', 'N', 
            'S', 'T', 'V', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH']
    
    colunasLarge = ['C', 'E', 'G', 'I', 'J', 'O', 'P', 'R', 'U', 'W', 
                    'F', 'Q', 'K', 'L', 'R', 'X', 'Y']
    
    worksheet.column_dimensions['A'].width = 13
    for i in range(len(colunasLarge)):
        worksheet.column_dimensions[colunasLarge[i]].width = 30
    for i in range(len(colunas)):
        worksheet.column_dimensions[colunas[i]].width = 20

if __name__ == '__main__':
    main()
