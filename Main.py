from datetime import datetime
import Requests_API
import Chamada_Relatorios
import Sheets_Excel

# SLA

data_atual = datetime.now()
dia_atual = data_atual.strftime('%d')
mes_atual = data_atual.strftime("%m")
ano_atual = data_atual.year

folder_path = input("Insira o diretório de export para os arquivos: ")

caminho_relatorio = f'{folder_path}RelatórioSLA_{dia_atual}-{mes_atual}-{ano_atual}.xlsx'

data_frame_SLA = Requests_API.recebimento_apiJira_SLA("sua empresa")

periodo_planilha = ['relatório-SLA_Semanal', 'relatório-SLA_Mensal']

dfSLA_Mensal = Chamada_Relatorios.relatorio_SLA(data_frame_SLA, periodo_planilha[1])
dfSLA_Semanal = Chamada_Relatorios.relatorio_SLA(data_frame_SLA, periodo_planilha[0])

df_Finais = [dfSLA_Semanal, dfSLA_Mensal]

Sheets_Excel.exportacao_xlsx(df_Finais, periodo_planilha, caminho_relatorio, 0)

# Status

data_frames_Status = Requests_API.recebimento_apiJira_Status('sua empresa')
planilha_nome = ['Resolvidos', 'Resolvidos-e-Abertos']

df_pivot_resolvidos = Chamada_Relatorios.relatorio_Status(data_frames_Status[0], planilha_nome[0])
df_pivot_resolvidos_abertos = Chamada_Relatorios.relatorio_Status(data_frames_Status[1], planilha_nome[1])

caminho_relatorio = f'{folder_path}RelatórioStatus_{dia_atual}-{mes_atual}-{ano_atual}.xlsx'

data_frames = [df_pivot_resolvidos, df_pivot_resolvidos_abertos]
Sheets_Excel.exportacao_xlsx(data_frames, planilha_nome, caminho_relatorio, 1)