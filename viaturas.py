# import libraries
import pandas as pd
import openpyxl
import win32com.client
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tabulate import tabulate


# read data from excel file
try:
    
    data = pd.read_excel('viaturas.xlsx')

except:
    
    print("Erro ao abrir ficheiro excel!")  

# read email list
try:
    
    fich = open('lista_emails.txt', 'r', encoding='UTF-8-SIG')
    recipients = fich.readlines()
    fich.close()

except:
    
    print("Erro ao abrir ficheiro texto!")

# calculate time diferences
def calculate_time(data, f):
    return relativedelta(current_dateTime, datetime.strptime(data, f))

# verify IPO dates
def passageiros(dif, row):
    if dif.years==3 or dif.years==5 or dif.years>=7:
        limit_date = (row['DATA MATRICULA'] + relativedelta(years=dif.years+1)).strftime('%Y-%m-%d')
        return [row['MARCA'], row['MODELO'], row['MATRICULA'], limit_date]

def mercadorias(dif, row):
    if dif.years>=1:
        limit_date = (row['DATA MATRICULA'] + relativedelta(years=dif.years+1)).strftime('%Y-%m-%d')
        return [row['MARCA'], row['MODELO'], row['MATRICULA'], limit_date]

# append emails
def app_email(row):
    if (row['EMAIL'] not in recipients):
        recipients.append(row['EMAIL'])
    
# send emails
def send_email(viat_html_table, subject, text, recipients):
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    #mail.SentOnBehalfOfName = 'geral@guimadiesel.pt'

    for recipient in recipients:
        mail.Recipients.Add(recipient).Type = 1
        
    mail.Subject = subject

    table_style = """
    <style>
    .gmail-table {
        border: solid 2px #81CDEB;
        border-collapse: collapse;
        border-spacing: 0;
        font: normal 14px Roboto, sans-serif;
    }

    .gmail-table thead th {
        background-color: #ACD7E8;
        border: solid 1px #81CDEB;
        color: #2D5F73;
        padding: 10px;
        text-align: left;
        text-shadow: 1px 1px 1px #fff;
    }

    .gmail-table tbody td {
        border: solid 1px #81CDEB;
        color: #333;
        padding: 10px;
        text-shadow: 1px 1px 1px #fff;
    }
    </style>
    """

    html_body = f"""
        <html>
        <head>
        {table_style}
        </head>
        <body>
            <p>Bom dia,<br>
            {text}<br>
            </p>
        {viat_html_table}
        </body>
        <footer>
            <p style="color:#999;font-size:8px;">© 2024 - Fernando Cunha</p>
        </footer>
        </html>
        """

    mail.HTMLBody = html_body

    mail.Send()


### MAIN ###
f = '%Y-%m-%d %H:%M:%S'   
current_dateTime = datetime.now().date()

viat_ipo_table = []
viat_rev_table = []

for index, row in data.iterrows():
    
    dataMat = str(row['DATA MATRICULA'])
    dataRev = str(row['DATA REVISAO'])
    dif_ipo = calculate_time(dataMat, f)
    dif_rev = calculate_time(dataRev, f)

    if dif_ipo.months==11 and dif_ipo.days>15:

        if (row['CATEGORIA'])=='Passageiros':
            viat_ipo_table.append(passageiros(dif_ipo, row))
            app_email(row)
                    
        elif (row['CATEGORIA'])=='Mercadorias':
                viat_ipo_table.append(mercadorias(dif_ipo, row))
                app_email(row)
                    
        else:
            print("Categoria inválida!")

    if dif_rev.months==0:     
        viat_rev_table.append([row['MARCA'], row['MODELO'], row['MATRICULA']])
        app_email(row)

# cleanup recipients
recipients = [x for x in recipients if str(x) != 'nan']       

# Notify
if len(viat_ipo_table) > 0:
    subject = "AVISO - Viaturas com data limite de inspeção próximas"
    text = "Segue informação sobre a(s) viatura(s) com datas limite de inspeção próximas:"
    viat_html_table = tabulate(viat_ipo_table, headers=["Marca", "Modelo", "Matricula", "Data Limite"],tablefmt='html')\
        .replace("<table>",'''<table class="gmail-table">''')        
    send_email(viat_html_table, subject, text, recipients)

if len(viat_rev_table) > 0:
    subject = "AVISO - Viaturas em período de revisão anual"
    text = "Segue informação sobre a(s) viatura(s) em período de revisão anual:"
    viat_html_table = tabulate(viat_rev_table, headers=["Marca", "Modelo", "Matricula"],tablefmt='html')\
        .replace("<table>",'''<table class="gmail-table">''')                
    send_email(viat_html_table, subject, text, recipients)


