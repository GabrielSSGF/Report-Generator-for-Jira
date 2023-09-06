import datetime as dt
from datetime import datetime, timedelta
from dateutil.parser import parse
import json
import logging
import openpyxl
import pandas as pd
import pytz
import requests
from requests.auth import HTTPBasicAuth

def main():
    data_atual = datetime.now()
    dia_atual = data_atual.strftime('%d')
    mes_atual = data_atual.strftime("%m")
    ano_atual = data_atual.year

    global dadosJson
    dadosJson = getDataFromJsonFile()

    logging.basicConfig(filename=f'{dadosJson["FilePathStatusLogs"]}_{dia_atual}-{mes_atual}-{ano_atual}.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    responseAberto, responseConcluido = getJiraAPIResponse(dadosJson)
    
    dataFrameAberto = responseJsonToDataframe(responseAberto, JQL_ABERTO)
    dataFrameConcluido = responseJsonToDataframe(responseConcluido, JQL_CONCLUIDO)
    
    planilhaNome = ['Resolvidos', 'Resolvidos-e-Abertos']

    dataFramePivotResolvidos = createStatusDataFrame(dataFrameConcluido, planilhaNome[0])
    dataFramePivotAbertosEResolvidos = createStatusDataFrame(dataFrameAberto, planilhaNome[1])

    caminhoRelatorio = f'{dadosJson["FilePathStatus"]}_{dia_atual}-{mes_atual}-{ano_atual}.xlsx'
    dataFrames = [dataFramePivotResolvidos, dataFramePivotAbertosEResolvidos]
    exportacaoXLSX(dataFrames, planilhaNome, caminhoRelatorio)

    logging.info("Relatório de Status gerado com sucesso")

def getDataFromJsonFile():
    with open('configData.json', 'r') as arquivoJson:
        dadosJson = json.load(arquivoJson)
    return dadosJson

def getJiraAPIResponse(configJson):
    try:
        url = f'https://{configJson["CompanyDomain"]}.atlassian.net/rest/api/3/search'
        
        global JQL_ABERTO, JQL_CONCLUIDO, AUTH
        
        JQL_ABERTO = configJson["JQLAbertoStatus"]
        JQL_CONCLUIDO = configJson["JQLConcluidoStatus"]
        AUTH = HTTPBasicAuth(configJson["Email"], configJson["APIToken"])

        headers = {
            "Accept": "application/json"
        }

        paramsAberto = {
            "jql": JQL_ABERTO,
            "maxResults": 100
        }

        paramsConcluido = {
            "jql": JQL_CONCLUIDO,
            "maxResults": 100
        }

        responseAberto = requests.get(url, headers=headers, auth=AUTH, params=paramsAberto)
        responseConcluido = requests.get(url, headers=headers, auth=AUTH, params=paramsConcluido)

        logging.info("*** Responses adquiridos com sucesso ***")
        
        return responseAberto, responseConcluido
   
    except Exception as e:
        logging.exception("Um erro ocorreu em recebimento_apiJira: %s", str(e))

def responseJsonToDataframe(response, jql):
    try:
        data = response.json()
        issues = data['issues']
        
        listaDados = []
        
        remainingTimeFirstResponse = None
        remaningTimeResolution = None
        

        for i in issues:
            key = i['key']
            cliente = i["fields"]["customfield_10163"]
            completedCyclesFirstResponse = i['fields']['customfield_10037']['completedCycles']
            completeCyclesResolution = i['fields']['customfield_10036']['completedCycles']
            
            remainingTimeFirstResponse = checkCompletedCycle(completedCyclesFirstResponse, 'customfield_10037', i)
            remaningTimeResolution = checkCompletedCycle(completeCyclesResolution, 'customfield_10036', i)
      
            row = [key, cliente, remainingTimeFirstResponse, remaningTimeResolution]
            
            listaDados.extend(getAdditionalIssueInfo(key, jql, row))
                        
        dataFrame = pd.DataFrame(listaDados, columns=["Key", "Cliente", "Time to first response",
                    "Time to resolution", "Status Transition.date", "Status Transition.to"])
        
        dataFrame.set_index("Key", inplace=True)
            
        logging.info("Json convertida para Data Frame.")
                
        return dataFrame
    except Exception as e:
        logging.exception("Um erro ocorreu em jsonToExcel: %s", str(e))
        
def checkCompletedCycle(completedCycle, customFieldType, loopIndex):            
    if not completedCycle:
        return getRemainingTimeIfCompleteCycleIsEmpty(customFieldType, loopIndex)
    else: return getRemainingTime(completedCycle)

def getRemainingTimeIfCompleteCycleIsEmpty(customField, loopIndex):
    ongoingCycle = None
    remainingTime = None
    ongoingCycle = loopIndex['fields'][customField].get('ongoingCycle')
    if ongoingCycle:
        remainingTime = ongoingCycle['remainingTime']['friendly']
        return remainingTime

def getRemainingTime(completedCyclesType):
    remainingTime = None
    for j in range(len(completedCyclesType)):
        if completedCyclesType[j]['remainingTime']['friendly'] is not None:
            remainingTime = completedCyclesType[j]['remainingTime']['friendly']
    return remainingTime
        
def getAdditionalIssueInfo(key, jql, row):
    listaDados = []
    requestAdicional = requestIssueExtraInfo(key, jql, AUTH)
    dadosTransitions = requestAdicional['values']
    createdDate = requestAdicional['values'][0]['created']
    
    quantidadeChamadas = 0
    
    for item in dadosTransitions:
        listaDados.extend(addItemsDosFieldsNaListaDados(item, createdDate, row, quantidadeChamadas))
        quantidadeChamadas += 1
        
    return listaDados
    
def addItemsDosFieldsNaListaDados(item, createdDate, row, quantidadeChamadas):
    listaDados = []
    for x in item['items']:
        if 'field' not in x or x["field"] != "status":
            continue
        
        rowCopy = row.copy()
        if quantidadeChamadas == 0:
            statusTransitionTo = x['fromString']
            rowCopy += [createdDate, statusTransitionTo]
        elif 'toString' in x:
            statusTransitionTo = x['toString']
            statusTransitionDate = item['created']
            rowCopy += [statusTransitionDate, statusTransitionTo]
        listaDados.append(rowCopy)
    return listaDados

def requestIssueExtraInfo(key, jql, auth):
    try:
        url = f'https://{dadosJson["CompanyDomain"]}.atlassian.net/rest/api/2/issue/{key}/changelog'
        headers = {
            "Accept": "application/json"
        }
        params = {
            "jql": jql,
            "maxResults": 100
        }
        response = requests.get(url, headers=headers, auth=auth, params=params)
        logging.info("Request realizado")
        return response.json()
    except Exception as e:
        logging.exception("Um erro ocorreu em request_issues: %s", str(e))

def createStatusDataFrame(dataFrame, planilha):
    try:
        dataFrame['Status Transition.date'] = dataFrame['Status Transition.date'].apply(lambda x: x if pd.notnull(x) else '')
        dataFrame['next_transition_date'] = dataFrame.groupby('Key')['Status Transition.date'].shift(-1)

        datetimeAtual = datetime.now(pytz.timezone("America/Sao_Paulo"))
        dataAtual = datetimeAtual.strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3] + datetimeAtual.strftime("%z")

        dataFrame.loc[dataFrame['Status Transition.to'] != 'Concluído', 'next_transition_date'] = dataFrame.loc[dataFrame['Status Transition.to'] != 'Concluído', 'next_transition_date'].fillna(dataAtual)

        dataFrame['Status Transition.date'] = dataFrame['Status Transition.date'].apply(lambda x: parse(x, fuzzy=True) if pd.notnull(x) else x)
        dataFrame['next_transition_date'] = dataFrame['next_transition_date'].apply(lambda x: parse(x, fuzzy=True) if pd.notnull(x) else x)

        dataFrame['Status Transition.date'] = dataFrame['Status Transition.date'].apply(lambda x: x.replace(tzinfo=None) if pd.notnull(x) else x)
        dataFrame['next_transition_date'] = dataFrame['next_transition_date'].apply(lambda x: x.replace(tzinfo=None) if pd.notnull(x) else x)

        dataFrame['Time Interval'] = dataFrame.apply(lambda row: subtrairDatas(row['Status Transition.date'], row['next_transition_date']), axis=1)

        valoresTotais = ['Total', 'Total Empresa', 'Total Clientes']
        total = calculoDeTotaleSLAs(dataFrame)

        dataFramePivot = dataFrame.pivot_table(index=['Key', 'Cliente', 
                                        'Time to first response', 'Time to resolution'],
                            columns='Status Transition.to',
                            values='Time Interval',
                            aggfunc='sum')

        for i in range(len(valoresTotais)):
            dataFramePivot[valoresTotais[i]] = total[i]

        colunas_numericas = dataFramePivot.select_dtypes(include='number')
        medias = colunas_numericas.mean()
        dataFramePivot = dataFramePivot.applymap(converterParaHorasMinutos)

        dataFramePivot = dataFramePivot.reset_index()

        novaLinha = pd.Series(medias, name='Média')
        
        # SLAS
        dataFramePivot['SLA First Response'] = dataFramePivot['Time to first response'].apply(lambda x: "Quebrado" if pd.notnull(x) and str(x).startswith('-') else "Cumprido")
        dataFramePivot['SLA Resolution'] = dataFramePivot['Time to resolution'].apply(lambda x: "Quebrado" if pd.notnull(x) and str(x).startswith('-') else "Cumprido")
        colunasTimeTo = ['Time to resolution', 'Time to first response']
        dataFramePivot = dataFramePivot.drop(colunasTimeTo, axis=1)

        dataFramePivot = dataFramePivot._append(novaLinha)
        dataFramePivot.loc['Média'] = dataFramePivot.loc['Média'].apply(converterParaHorasMinutos)
        
        dataFramePivot = dataFramePivot.set_index('Key').sort_values(by='Key', ascending=True)
        dataFramePivot = dataFramePivot.reindex(dataFramePivot.index[1:].tolist() + [dataFramePivot.index[0]])
        dataFramePivot.rename(index={"-": "Média"}, inplace=True)

        if "Aguardando suporte" in dataFramePivot.columns:
            dataFramePivot = dataFramePivot.drop("Aguardando suporte", axis=1)

        if "Aberto" in dataFramePivot.columns:
            dataFramePivot = dataFramePivot.drop("Aberto", axis=1)

        dataFramePivot = dataFramePivot.style.apply(estiloNegrito, dataFrame=dataFramePivot, axis=1)

        logging.info(f"*** DataFrame '{planilha}' de Status criado ***")

        return dataFramePivot
    except Exception as e:
        logging.exception("Um erro ocorreu em createDataFramePivot: %s", str(e))

def subtrairDatas(dataInicio, dataFim):
    try:
        if pd.isnull(dataInicio) or pd.isnull(dataFim):
            return None
        
        
        dataInicio = adjustDataInicioFimToServiceTime(dataInicio)
        dataFim = adjustDataInicioFimToServiceTime(dataFim)
       
        segundosTotais = (dataFim - dataInicio).total_seconds()
        
        if dataInicio.time() > dataFim.time():
            dataFim = dataFim + dt.timedelta(days=1)

        diferenca = dataFim - dataInicio
        
        segundosTotais = removeSegundosForaDoServiceTime(segundosTotais, diferenca, dataInicio)
        
        logging.info("Datas subtraídas respeitando o horário comercial")

        return segundosTotais
    except Exception as e:
        logging.exception("Um erro ocorreu em subtrair_datas: %s", str(e))
        
def adjustDataInicioFimToServiceTime(dataTipo):
    sabado = 5
    domingo = 6
    diaSemana = dataTipo.weekday()
    
    if diaSemana == sabado:
        dataTipo += timedelta(days=2)
        dataTipo = dataTipo.replace(hour=9, minute=0)

    elif diaSemana == domingo:
        dataTipo += timedelta(days=1)
        dataTipo = dataTipo.replace(hour=9, minute=0)

    if dataTipo.time().hour > 18:
        dataTipo = dataTipo.replace(hour=18, minute=0)

    elif dataTipo.time().hour < 9:
        dataTipo = dataTipo.replace(hour=9, minute=0)
    
    return dataTipo

def removeSegundosForaDoServiceTime(segundosTotais, diferenca, dataInicio):
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
        segundosTotais -= 54000
        
    elif diasUteis > 2:
        segundosForaServico = (diasUteis-2) * 54000 + 54000
        segundosTotais -= segundosForaServico
    
    return segundosTotais

def calculoDeTotaleSLAs(dataFrame):
    try:
        statusCliente = [dadosJson["StatusCustomer"]]
        
        statusEmpresa = [dadosJson["StatusCompany"]]

        total = dataFrame.pivot_table(index=['Key', 'Cliente'],
                    columns='Status Transition.to',
                    values='Time Interval',
                    aggfunc='sum').sum(axis=1)
        
        totalCliente = totalPorStatus(statusCliente, dataFrame)
        totalEmpresa = totalPorStatus(statusEmpresa, dataFrame)
        
        logging.info("*** Total de tempo por grupos calculado ***")

        return total, totalEmpresa, totalCliente
    except Exception as e:
        logging.exception("Um erro ocorreu em calculoDeTotaleSLAs: %s", str(e))
        
def totalPorStatus(status, dataFrame):
    try:
        total = dataFrame[dataFrame['Status Transition.to'].isin(status)].pivot_table(
            index=['Key', 'Cliente'],
            columns='Status Transition.to',
            values='Time Interval',
            aggfunc='sum'
        ).sum(axis=1)
        
        logging.info("Total de grupo definido")
        
        return total
    except Exception as e:
        logging.exception("Um erro ocorreu em totalPorStatus: %s", str(e))
        
def converterParaHorasMinutos(valor):
    try:
        if pd.isnull(valor):
            return "-"

        hora = int(valor / 3600)
        minutos = int(valor % 3600) // 60
        
        logging.info("Valores convertidos para horas e minutos")

        return f'{hora:02d}h{minutos:02d}'
    except Exception as e:
        logging.exception("Um erro ocorreu em converterParaHorasMinutos: %s", str(e))

def estiloNegrito(row, dataFrame):
    try:
        if row.name == dataFrame.index[-1]:
            return ['font-weight: bold'] * len(row)
        
        logging.info("Estilo da linha definido em negrito")
        
        return [''] * len(row)
    except Exception as e:
        logging.exception("Um erro ocorreu em estilo_negrito: %s", str(e))

def exportacaoXLSX(data_frames, planilhas, caminho_relatorio):
    try:
        writer = pd.ExcelWriter(caminho_relatorio, engine='openpyxl')
        for i, dataFrame in enumerate(data_frames):
            dataFrame.to_excel(writer, sheet_name=planilhas[i])
            workbook = writer.book
            worksheet = workbook[planilhas[i]]
            padraoDeColunas(worksheet)
        writer._save()
        
        logging.info("DataFrames de Status exportado como Excel")
        
    except Exception as e:
        logging.exception("Um erro ocorreu em exportacao_xlsx: %s", str(e))

def padraoDeColunas(worksheet):
    try:
        colunas = ['D', 'H', 'M', 'N', 
                'S', 'T', 'V', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']
        
        colunas_large = ['C', 'E', 'G', 'I', 'J', 'O', 'P', 'R', 'U', 'W', 
                        'F', 'Q', 'K', 'L', 'R', 'X', 'Y']
        
        # Definir a largura das colunas
        worksheet.column_dimensions['A'].width = 13
        worksheet.column_dimensions['B'].width = 62

        for i in range(len(colunas_large)):
            worksheet.column_dimensions[colunas_large[i]].width = 30

        for i in range(len(colunas)):
            worksheet.column_dimensions[colunas[i]].width = 20
            
        logging.info("Padrão de colunas estabelecido")
        
    except Exception as e:
        logging.exception("Um erro ocorreu em padrao_de_colunas: %s", str(e))

if __name__ == '__main__':
    main()
