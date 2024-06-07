### IMPORT LIBRARIES ###

import pandas as pd
import win32com.client
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tabulate import tabulate
import os
from dotenv import load_dotenv
import pymongo


### READ DATA ###

# Load environment variables
load_dotenv()
try:
    ENV_USER = os.getenv('USER')
    ENV_PASSWORD = os.getenv('PASSWORD')
except KeyError:
    raise KeyError("Token not available!")

# Initialize the 'client' variable to None
client = None

# Retrieve data from MongoDB database
try:
    # Creating a MongoClient to connect to the local MongoDB server
    client = pymongo.MongoClient(f'mongodb+srv://{ENV_USER}:{ENV_PASSWORD}@cluster0.gmd3vkc.mongodb.net/guimadiesel?retryWrites=true&w=majority')
    # Getting the 'guimadiesel' database from the MongoDB server
    db = client['guimadiesel']
    # Getting the 'cars' collection from the 'mongodb' database
    collection = db['cars']
    print("MongoDB connection established successfully.")
except Exception as e:
    # Handling exceptions and printing an error message if connection fails
    print(f"Error connecting to MongoDB: {e}")
finally:
    # Close the MongoDB client if it was initialized
    if client is None:
        client.close()
        print("Connection closed.")

# Creating a Pandas DataFrame from collection
data = pd.DataFrame(list(collection.find()))
data['dataMat']= pd.to_datetime(data['data_mat'])
data['dataRev']= pd.to_datetime(data['data_rev'])

# Read email list
try:   
    fich = open('lista_emails.txt', 'r', encoding='UTF-8-SIG')
    recipients = fich.readlines()
    fich.close()
except:   
    print("Error opening emails file!")


### FUNCTIONS ###

# Calculate time diferences
def calculate_time(data, f):
    return relativedelta(current_dateTime, datetime.strptime(data, f))

# Verify IPO dates
def passageiros(dif, row):
    if dif.years==3 or dif.years==5 or dif.years>=7:
        limit_date = (row['dataMat'] + relativedelta(years=dif.years+1)).strftime('%Y-%m-%d')
        return [row['marca'], row['modelo'], row['matricula'], limit_date]

def mercadorias(dif, row):
    if dif.years>=1:
        limit_date = (row['dataMat'] + relativedelta(years=dif.years+1)).strftime('%Y-%m-%d')
        return [row['marca'], row['modelo'], row['matricula'], limit_date]

# Append email recipients
def app_email(row):
    if (row['email'] not in recipients):
        recipients.append(row['email'])
    
# Send emails
def send_email(viat_html_table, subject, text, recipients):
    
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = 'geral@guimadiesel.pt'
    mail.Subject = subject

    for recipient in recipients:
        mail.Recipients.Add(recipient).Type = 1
     
# Message table styling
    table_style = """
    <style>
    .outlook-table {
        border: solid 2px #81CDEB;
        border-collapse: collapse;
        border-spacing: 0;
        font: normal 14px Roboto, sans-serif;
    }
    .outlook-table thead th {
        background-color: #ACD7E8;
        border: solid 1px #81CDEB;
        color: #2D5F73;
        padding: 10px;
        text-align: left;
        text-shadow: 1px 1px 1px #fff;
    }
    .outlook-table tbody td {
        border: solid 1px #81CDEB;
        color: #333;
        padding: 10px;
        text-shadow: 1px 1px 1px #fff;
    }
    </style>
    """

# Message composing
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

# Message sending
    mail.Send()


### MAIN ###

# Initializing variables
f = '%Y-%m-%d %H:%M:%S'   
current_dateTime = datetime.now().date()

viat_ipo_table = []
viat_rev_table = []

# Iterate data values
for index, row in data.iterrows():
    
    dataMat = str(row['dataMat'])
    dataRev = str(row['dataRev'])
    dif_ipo = calculate_time(dataMat, f)
    dif_rev = calculate_time(dataRev, f)

    if dif_ipo.months==11 and dif_ipo.days>15:
        if (row['categ'])=='Passageiros':
            viat_ipo_table.append(passageiros(dif_ipo, row))
            app_email(row)                   
        elif (row['categ'])=='Mercadorias':
                viat_ipo_table.append(mercadorias(dif_ipo, row))
                app_email(row)                   
        else:
            print("Categoria inválida!")

    if dif_rev.months==0:
        viat_rev_table.append([row['marca'], row['modelo'], row['matricula']])
        app_email(row)

# Cleanup email recipients
recipients = [x for x in recipients if str(x) != 'nan']       

# Notify recipients
if len(viat_ipo_table) > 0:
    subject = "AVISO - Viaturas com data limite de inspeção próximas"
    text = "Segue informação sobre a(s) viatura(s) com datas limite de inspeção próximas:"
    viat_html_table = tabulate(viat_ipo_table, headers=["Marca", "Modelo", "Matricula", "Data Limite"],tablefmt='html')\
        .replace("<table>",'''<table class="outlook-table">''')        
    send_email(viat_html_table, subject, text, recipients)

if len(viat_rev_table) > 0:
    subject = "AVISO - Viaturas em período de revisão anual"
    text = "Segue informação sobre a(s) viatura(s) em período de revisão anual:"
    viat_html_table = tabulate(viat_rev_table, headers=["Marca", "Modelo", "Matricula"],tablefmt='html')\
        .replace("<table>",'''<table class="outlook-table">''')                
    send_email(viat_html_table, subject, text, recipients)
