# import libraries
import pandas as pd
#import openpyxl
import win32com.client
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tabulate import tabulate


# read data from excel file
try:
    
    data = pd.read_excel('viaturas.xlsx')

except:
    
    print("Erro ao abrir ficheiro!")  

# read email list
try:
    
    fich = open('lista_emails.txt', 'r', encoding='UTF-8-SIG')
    recipients = fich.readlines()
    fich.close()

except:
    
    print("Erro ao abrir ficheiro!")

# calculate time diferences
def calculate_time(data, f):
    return relativedelta(current_dateTime, datetime.strptime(data, f))

# verify IPO dates
def passageiros(dif):
    if dif.years==4 or dif.years==4 or dif.years>=8:
        return True

def mercadorias(dif):
    if dif.years>=2:
        return True

# send emails
def send_email(viat_html_table, subject, text, recipients):
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.SendUsingAccount = 'geral@guimadiesel.pt'

    for recipient in recipients:
        mail.Recipients.Add(recipient).Type = 1
        
    mail.Subject = subject

    table_style = """
    <style>
    .gmail-table {
        border: solid 2px #DDEEEE;
        border-collapse: collapse;
        border-spacing: 0;
        font: normal 14px Roboto, sans-serif;
    }

    .gmail-table thead th {
        background-color: #DDEFEF;
        border: solid 1px #DDEEEE;
        color: #336B6B;
        padding: 10px;
        text-align: left;
        text-shadow: 1px 1px 1px #fff;
    }

    .gmail-table tbody td {
        border: solid 1px #DDEEEE;
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
        </html>
        """

    mail.HTMLBody = html_body

    mail.Send()


### MAIN ###
f = '%Y-%m-%d %H:%M:%S'
current_dateTime = datetime.now()

viat_ipo_table = []
viat_rev_table = []

for index, row in data.iterrows():
    
    dataMat = str(row['DATA MATRICULA'])
    dataRev = str(row['DATA REVISAO'])
    dif_ipo = calculate_time(dataMat, f)
    dif_rev = calculate_time(dataRev, f)

    if dif_ipo.months==11 and dif_ipo.days<16:

        if (row['CATEGORIA'])=='Passageiros':
            if passageiros(dif_ipo):
                viat_ipo_table.append([row['MARCA'], row['MODELO'], row['MATRICULA'], row['DATA MATRICULA']])
                if (row['EMAIL'] not in recipients):
                    recipients.append(row['EMAIL'])

        elif (row['CATEGORIA'])=='Mercadorias':
            if mercadorias(dif_ipo):
                viat_ipo_table.append([row['MARCA'], row['MODELO'], row['MATRICULA'], row['DATA MATRICULA']])
                if (row['EMAIL'] not in recipients):
                    recipients.append(row['EMAIL'])

        else:
            print("Categoria inválida!")

    if dif_rev.months==0:
        
        viat_rev_table.append([row['MARCA'], row['MODELO'], row['MATRICULA']])
        if ((row['EMAIL']) not in recipients):
            recipients.append(row['EMAIL'])

recipients = [x for x in recipients if str(x) != 'nan']
            
print(viat_ipo_table)
print(viat_rev_table)
print(recipients)

# Notify
if len(viat_ipo_table) > 0:
    subject = "AVISO - Viaturas com data limite de inspeção em breve"
    text = "Aqui estão a(s) viatura(s) com datas limite de inspeção próximas:"
    viat_html_table = tabulate(viat_ipo_table, headers=["Marca", "Modelo", "Matricula", "Data Limite"],tablefmt='html')\
        .replace("<table>",'''<table class="gmail-table">''')        
    #send_email(viat_html_table, subject, text, recipients)
    print(viat_html_table)

if len(viat_rev_table) > 0:
    subject = "AVISO - Viaturas em período de revisão anual"
    text = "Aqui estão a(s) viatura(s) em período de revisão anual:"
    viat_html_table = tabulate(viat_rev_table, headers=["Marca", "Modelo", "Matricula"],tablefmt='html')\
        .replace("<table>",'''<table class="gmail-table">''')                
    #send_email(viat_html_table, subject, text, recipients)
    print(viat_html_table)


