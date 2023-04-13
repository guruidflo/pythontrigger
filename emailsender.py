import smtplib
import json
from email.message import EmailMessage
import datetime
from datetime import datetime
from datetime import timedelta
from datetime import date
import pandas as pd
import openpyxl
from openpyxl import Workbook

today = date.today()
date_str = datetime.now().strftime('%Y-%m-%d')
email_address = "fabtrakr@ideplexports.com"
email_password = "%iDflo@777%"
xlsx_file_name =  'file-'+date_str+'.xlsx'
json_file_name = 'file-'+date_str+'.json'

with open(json_file_name, 'r') as file:
 load_json = json.load(file)

TotalProduction = 0
Totalmachines = 0
sewingSamProduced = 0
sewingMachineMinutes = 0
globalSamProduced = 0
globalMachineMinutes = 0

for x in load_json:
  TotalProduction += x['outputQuantity'] 
  Totalmachines += x['numberOfMachines']
  sewingSamProduced += x['sewingSamProduced']
  sewingMachineMinutes += x['sewingMachineMinutes']
  globalSamProduced += x['globalSamProduced']
  globalMachineMinutes += x['globalMachineMinutes']

sewingefficinecy = (sewingSamProduced/sewingMachineMinutes)*100
globalefficiency = (globalSamProduced/globalMachineMinutes)*100

sewingeff = format(sewingefficinecy, '.2f')
globaleff = format(globalefficiency, '.2f')

msg = EmailMessage()
msg['Subject'] = "Overview Report : " +date_str
msg['From'] = email_address
# msg['To'] = "gurumurthy@indian-designs.net"
msg['To'] = "gurumurthy@indian-designs.net,madhukar@indian-designs.com"
msg['Cc'] = "gurumurthy@indian-designs.com"
msg.set_content("""Dear Team,

Please find the below Attached Overview report and Mentioned Summary for """+str(date_str)+""".

Summary :- 

➜ Total Production = """+str(TotalProduction)+"""
➜ Total Machines Running = """+str(Totalmachines)+"""
➜ Sewing Efficiency = """+str(sewingeff)+"""%
➜ Global Efficiency = """+str(globaleff)+"""%

Regards,
Team IDFlo
This is an Auto Generated Email, Please do not reply.""")

with open(xlsx_file_name,"rb") as f:
    file_data=f.read()
    file_name=f.name
    print(file_name)
    msg.add_attachment(file_data,maintype="application",subtype="xlsx",filename=file_name)
    
with smtplib.SMTP_SSL('smtp3.netcore.co.in', 465) as smtp:
    smtp.login(email_address, email_password)
    smtp.send_message(msg)
