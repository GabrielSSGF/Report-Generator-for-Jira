import pandas as pd
import datetime
import requests
import json
from requests.auth import HTTPBasicAuth
from datetime import date, datetime
import logging

def main():
    dataAtual = datetime.now()
    diaAtual = dataAtual.strftime('%d')
    mesAtual = dataAtual.strftime("%m")
    anoAtual = dataAtual.year
    
    dadosJson = getDataFromJsonFile()

    logging.basicConfig(filename=f'{dadosJson["FilePathSLALogs"]}_{diaAtual}-{mesAtual}-{anoAtual}.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    caminhoRelatorio = f'{dadosJson["FilePathSLA"]}_{diaAtual}-{mesAtual}-{anoAtual}.xlsx'
    
    periodoPlanilha = ['relatório-SLA_Semanal', 'relatório-SLA_Mensal']
    
    dataFrame = jsonToDataFrame(getApiJsonSLA())
    
    dataFrameSLAMensal = createSLADataFrame(dataFrame, periodoPlanilha[1])
    dataFrameSLASemanal = createSLADataFrame(dataFrame, periodoPlanilha[0])

    dataFramesFinais = [dataFrameSLASemanal, dataFrameSLAMensal]

    exportacaoXLSX(dataFramesFinais, periodoPlanilha, caminhoRelatorio)

    logging.info("Relatório de SLA gerado com sucesso")
    
def getDataFromJsonFile():
    with open('configData.json', 'r') as arquivoJson:
            dadosJson = json.load(arquivoJson)
    return dadosJson

def getApiJsonSLA(configJson):
    try:
        url = f'https://{configJson["CompanyDomain"]}.atlassian.net/rest/api/3/search'
        jql = configJson["JQLSLA"]
        auth = HTTPBasicAuth(configJson["Email"], configJson["APIToken"])

        headers = {
            "Accept": "application/json"
        }

        params = {
            "jql": jql,
            "maxResults": 100
        }

        response = requests.get(url, headers=headers, auth=auth, params=params)

        if response.status_code != 200:
            logging.error("API request error: %d - %s", response.status_code, response.text)
            return
        
        return response
    
    except Exception as e:
        logging.exception("Um erro ocorreu getApiJsonSLA: %s", str(e))
    
def jsonToDataFrame(response):
    try:
        listaDados = []
        
        data = response.json()
        issues = data["issues"]

        for issue in issues:
            key = issue["key"]
            created = issue["fields"]["created"]
            completedCyclesFirstResponse = issue['fields']['customfield_10036']['completedCycles']
            completedCyclesResolution = issue['fields']['customfield_10037']['completedCycles']
            remainingTimeFirstResponse = completedCyclesFirstResponse[-1]['remainingTime']['friendly'] if completedCyclesFirstResponse else None
            remainingTimeResolution = completedCyclesResolution[-1]['remainingTime']['friendly'] if completedCyclesResolution else None
            resolutionDate = issue["fields"]["resolutiondate"]

            row = [key, created, remainingTimeFirstResponse, remainingTimeResolution, resolutionDate]
            listaDados.append(row)
            
        print(listaDados)
        
        dataFrame = pd.DataFrame(listaDados, columns=["Key", "Created", "Time to resolution", "Time to first response", "Resolved"])
        logging.info("Json convertida para Data Frame.")
        return dataFrame
    
    except Exception as e:
        logging.exception("Um erro ocorreu jsonToDataFrame: %s", str(e))

def createSLADataFrame(dataFrame, nomeDataFrame):
    try:
        tempo = dataFrame[['Time to first response', 'Time to resolution']].copy()

        tempo['Time to first response'] = pd.to_timedelta(tempo['Time to first response'])
        tempo['Time to resolution'] = pd.to_timedelta(tempo['Time to resolution'])

        dataFrame['Created'] = pd.to_datetime(dataFrame['Created'], format='%Y-%m-%dT%H:%M:%S.%f%z')
        dataFrame['Resolved'] = pd.to_datetime(dataFrame['Resolved'], format='%Y-%m-%dT%H:%M:%S.%f%z')

        nomeDataFrame = nomeDataFrame.split('_')

        if nomeDataFrame[-1] == "Mensal":
            periodo = "Month"
            dataFrame_nova_tabela = tempoRelatorio(periodo, dataFrame, tempo)
            return dataFrame_nova_tabela 

        periodo = 'Week'
        dataFrameNovaTabela = tempoRelatorio(periodo, dataFrame, tempo)
        
        logging.info(f"DataFrame 'SLA-{nomeDataFrame[-1]}' criado")

        return dataFrameNovaTabela

    except Exception as e:
        logging.exception("Um erro ocorreu em createSLADataFrame: %s", str(e))

def tempoRelatorio(periodo, dataFrame, tempo):
    try:
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
        
        logging.info("Tabela de tempo criada")
            
        return dataFrameNovaTabela
    
    except Exception as e:
        logging.exception("Um erro ocorreu em tempoRelatorio: %s", str(e))

def calculoDePorcentagens(tempo, dataFrame, qntd, periodo):
    try:
        slaFirstResponse = (tempo[tempo['Time to first response'] > pd.Timedelta(0)].groupby(dataFrame[periodo]).size() / qntd * 100).map(lambda x: round(x) if pd.notnull(x) else "None")
        slaResolution = (tempo[tempo['Time to resolution'] > pd.Timedelta(0)].groupby(dataFrame[periodo]).size() / qntd * 100).map(lambda x: round(x) if pd.notnull(x) else "None")
        
        logging.info("Porcentagens calculadas")
        
        return slaFirstResponse, slaResolution
    except Exception as e:
        logging.exception("Um erro ocorreu em calculoDePorcentagens: %s", str(e))

def criacaoTabelaSemanal(slas, qntd):
    try: 
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
        
        logging.info("Tabela semanal criada")
        
        return dataFrame
    
    except Exception as e:
        logging.exception("Um erro ocorreu em criacaoTabelaSemanal: %s", str(e))

def criacaoTabelaMensal(slas, qntd):
    try:
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
        
        logging.info("Tabela mensal criada")
        
        return dataFrame

    except Exception as e:
        logging.exception("Um erro ocorreu em criacaoTabelaMensal: %s", str(e))

def exportacaoXLSX(dataFrames, planilhas, caminhoRelatorio):
    try:
        writer = pd.ExcelWriter(caminhoRelatorio, engine='openpyxl')

        for i, dataFrame in enumerate(dataFrames):
            dataFrame.to_excel(writer, sheet_name=planilhas[i])
            workbook = writer.book
            worksheet = workbook[planilhas[i]]
            padraoDeColunas(worksheet)

        writer._save()
        logging.info("DataFrame de SLA exportado como Excel")
        
    except Exception as e:
        logging.exception("Um erro ocorreu em exportacaoXLSX: %s", str(e))

def padraoDeColunas(worksheet):
    try: 
        colunas = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'R', 
                'S', 'T', 'V', 'Z', 'AA', 'AB', 'AC', 'AD']
        
        worksheet.column_dimensions['A'].width = 22
        worksheet.column_dimensions['B'].width = 22

        for i in range(len(colunas)):
            worksheet.column_dimensions[colunas[i]].width = 20
            
        logging.info("Padrão de colunas estabelecido")

    except Exception as e:
        logging.exception("Um erro ocorreu em padraoDeColunas: %s", str(e))

if __name__ == '__main__':
    main()
