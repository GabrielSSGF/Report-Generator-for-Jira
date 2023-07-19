import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

"""
    Tudo relacionado a API do SLA
"""

def recebimento_apiJira_SLA(minha_empresa):
    url = f"https://{minha_empresa}.atlassian.net/rest/api/3/search"
    jql = 'project = "INSIRA SUA KEY" AND created >= startOfYear() AND resolution = Resolvido order by created DESC'
    auth = HTTPBasicAuth("INSIRA SEU E-MAIL", "INSIRA SEU TOKEN")

    headers = {
        "Accept": "application/json"
    }

    params = {
        "jql": jql,
        "maxResults": 100
    }

    response = requests.get(url, headers=headers, auth=auth, params=params)

    if response.status_code == 200:
        data = response.json()
        issues = data["issues"]

        rows = []

        for issue in issues:
            key = issue["key"]
            created = issue["fields"]["created"]
            completed_cycles_firstResponse = issue['fields']['customfield_10036']['completedCycles']
            completed_cycles_resolution = issue['fields']['customfield_10037']['completedCycles']
            remaining_time_firstResponse = completed_cycles_firstResponse[-1]['remainingTime']['friendly'] if completed_cycles_firstResponse else None
            remaining_time_resolution = completed_cycles_resolution[-1]['remainingTime']['friendly'] if completed_cycles_resolution else None
            resolution_date = issue["fields"]["resolutiondate"]

            row = [key, created, remaining_time_firstResponse, remaining_time_resolution, resolution_date]
            rows.append(row)

        headers = ["Key", "Created", "Time to resolution", "Time to first response", "Resolved"]
        df = pd.DataFrame(rows, columns=headers)

        return df

    else:
        print("Error:", response.status_code, response.text)

"""
    Tudo relacionado a API do Status
"""

def requests_issues(key, jql, auth, minha_empresa):
    url = f"https://{minha_empresa}.atlassian.net/rest/api/2/issue/{key}/changelog"
    headers = {
        "Accept": "application/json"
    }
    params = {
        "jql": jql,
        "maxResults": 100
    }
    response = requests.get(url, headers=headers, auth=auth, params=params)
    return response.json()

def jsonToDataFrame(params, url, auth, headers, jql, sheet):
    
    header_excel = ["Key", "Time to resolution",
                    "Time to first response", "Status Transition.date", "Status Transition.to"]
    sheet.append(header_excel)

    response = requests.get(url, headers=headers, auth=auth, params=params)
    data = response.json()
    issues = data['issues']

    ongoing_cycle = None
    remaining_time_firstResponse = None
    remaining_time_resolution = None

    for i in issues:
        key = i['key']
        completed_cycles_firstResponse = i['fields']['customfield_10036']['completedCycles']
        completed_cycles_resolution = i['fields']['customfield_10037']['completedCycles']

        for j in range(len(completed_cycles_firstResponse)):
            if completed_cycles_firstResponse[j]['remainingTime']['friendly'] is not None:
                remaining_time_firstResponse = completed_cycles_firstResponse[j]['remainingTime']['friendly']

        for n in range(len(completed_cycles_resolution)):
            if completed_cycles_resolution[n]['remainingTime']['friendly'] is not None:
                remaining_time_resolution = completed_cycles_resolution[n]['remainingTime']['friendly']

        if not completed_cycles_firstResponse:
            ongoing_cycle = i['fields']['customfield_10036'].get('ongoingCycle')
            if ongoing_cycle:
                remaining_time_firstResponse = ongoing_cycle['remainingTime']['friendly']

        if not completed_cycles_resolution:
            ongoing_cycle = i['fields']['customfield_10037'].get('ongoingCycle')
            if ongoing_cycle:
                remaining_time_resolution = ongoing_cycle['remainingTime']['friendly']

        request_atual = requests_issues(key, jql, auth)
        dados_transitions = request_atual['values']
        created_date = request_atual['values'][0]['created']
        y = 0
        for item in dados_transitions:
            for x in item['items']:
                if 'field' in x and x["field"] == "status":
                    if y == 0:
                        transition_to = x['fromString']
                        row = [key, remaining_time_resolution,
                               remaining_time_firstResponse, created_date, transition_to]
                        sheet.append(row)
                        transition_to = x['toString']
                        transition_date = item['created']
                        row = [key, remaining_time_resolution,
                               remaining_time_firstResponse, transition_date, transition_to]
                        sheet.append(row)      
                        y += 1
                        
                    elif 'toString' in x:
                        transition_to = x['toString']
                        transition_date = item['created']

                        row = [key, remaining_time_resolution,
                                remaining_time_firstResponse, transition_date, transition_to]
                        sheet.append(row)


def recebimento_apiJira_Status(minha_empresa):
    url_issues = f"https://{minha_empresa}.atlassian.net/rest/api/3/search"
    jql_concluido = 'project = "INSIRA SUA KEY" AND resolved >= startOfMonth() AND status != Cancelado'
    jql_aberto = 'project = "INSIRA SUA KEY" AND resolved >= startOfMonth() AND status != Cancelado OR project = "INSIRA SUA KEY" AND created >= startOfYear() AND status != Cancelado AND resolution = Unresolved'
    auth = HTTPBasicAuth("INSIRA SEU E-MAIL", "INSIRA SEU TOKEN")

    headers = {
        "Accept": "application/json"
    }

    params_aberto = {
        "jql": jql_aberto,
        "maxResults": 100
    }

    params_concluido = {
        "jql": jql_concluido,
        "maxResults": 100
    }

    df_fechada = jsonToDataFrame(params_concluido, url_issues, auth, headers, jql_concluido, minha_empresa)
    df_aberta = jsonToDataFrame(params_aberto, url_issues, auth, headers, jql_aberto, minha_empresa)
    
    return df_fechada, df_aberta