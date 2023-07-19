import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import datetime

"""
    REALIZAR O MERGE DAS FUNÇÕES EM UMA ÚNICA FUNÇÃO
"""

def enviar_emailSemanal():
    remetente = 'INSIRA REMETENTE'
    destinatarios = ['']

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = f'Relatório Semanal de Status e SLAs - {dia_atual}/{mes_atual}/{ano_atual}'

    corpo = """
    Bom dia!

    Segue em anexo o relatório semanal de status e SLA de cada ticket gerado no processo de atendimento.
    Nestes relatórios constam os períodos entre cada status individual, assim como os SLA's vencidos e resolvidos nesta(neste) semana(mês).

    Gerado por automação.

    Atenciosamente,

    [EMPRESA]
    """

    msg.attach(MIMEText(corpo, 'plain'))

    caminho_arquivo1 = f'/home/ubuntu/Relatorios/RelatórioSLA_{dia_atual}-{mes_atual}-{ano_atual}.xlsx'
    caminho_arquivo2 = f'/home/ubuntu/Relatorios/RelatórioStatus_{dia_atual}-{mes_atual}-{ano_atual}.xlsx'
    nome_arquivo = [f'RelatórioSLA.xlsx', f'RelatórioStatus_{dia_atual}-{mes_atual}-{ano_atual}.xlsx']
    arquivos = [caminho_arquivo1, caminho_arquivo2]
    
    for i, caminho_arquivo in enumerate(arquivos):
        with open(caminho_arquivo, 'rb') as anexo:
            part = MIMEApplication(anexo.read(), Name=nome_arquivo[i])
            part['Content-Disposition'] = f'attachment; filename="{nome_arquivo[i]}"'
            msg.attach(part)

    # SMTP.
    servidor = smtplib.SMTP('smtp.office365.com', 587)
    servidor.starttls()
    servidor.login(remetente, 'INSIRA SUA SENHA')
    texto = msg.as_string()
    servidor.sendmail(remetente, destinatarios, texto)
    servidor.quit()

def enviar_emailMensal():
    remetente = 'INSIRA SEU EMAIL'
    destinatarios = ['INSIRA', 'MULTIPLOS', 'DESTINATARIOS SE DESEJAR']

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ', '.join(destinatarios)
    msg['Subject'] = f'Relatório Mensal de Status e SLAs - {mes_ontem}/{ano_ontem}'

    corpo = """
    Bom dia!

    Segue em anexo o relatório mensal de status e SLA de cada ticket gerado no processo de atendimento.
    Nestes relatórios constam os períodos entre cada status individual, assim como os SLA's vencidos e resolvidos nesta(neste) semana(mês).

    Gerado por automação.

    Atenciosamente,

    [EMPRESA]
    """

    msg.attach(MIMEText(corpo, 'plain'))

    caminho_arquivo1 = f'/home/ubuntu/Relatorios/RelatórioSLA_{dia_ontem}-{mes_ontem}-{ano_ontem}.xlsx'
    caminho_arquivo2 = f'/home/ubuntu/Relatorios/RelatórioStatus_{dia_ontem}-{mes_ontem}-{ano_ontem}.xlsx'
    nome_arquivo = [f'RelatórioSLA.xlsx', f'RelatórioStatus_{mes_ontem}-{ano_ontem}.xlsx']
    arquivos = [caminho_arquivo1, caminho_arquivo2]
    
    for i, caminho_arquivo in enumerate(arquivos):
        with open(caminho_arquivo, 'rb') as anexo:
            part = MIMEApplication(anexo.read(), Name=nome_arquivo[i])
            part['Content-Disposition'] = f'attachment; filename="{nome_arquivo[i]}"'
            msg.attach(part)

    # SMTP.
    servidor = smtplib.SMTP('smtp.office365.com', 587)
    servidor.starttls()
    servidor.login(remetente, 'INSIRA SUA SENHA')
    texto = msg.as_string()
    servidor.sendmail(remetente, destinatarios, texto)
    servidor.quit()

data_atual = datetime.datetime.now()
dia_atual = data_atual.strftime('%d')
mes_atual = data_atual.strftime("%m")
ano_atual = data_atual.year

data_ontem = data_atual - datetime.timedelta(days=1)
dia_ontem = data_ontem.strftime('%d')
mes_ontem = data_ontem.strftime('%m')
ano_ontem = data_ontem.year