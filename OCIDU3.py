import requests
import json
import pandas as pd
from pathlib import Path
from datetime import datetime,timedelta
from datetime import date
from openpyxl import Workbook
from firebasetokengenerator import token
import os.path

date_today  = date.today() 
date_str = datetime.now().strftime('%Y-%m-%d')
date_input = datetime.now().strftime('%Y%m%d')
json_file_name = 'fileIDU3-'+date_str+'.json'
xlsx_file_name =  'fileIDU3-'+date_str+'.xlsx'
path = "C:\\Users\\Gurumurthy\\Documents\\Overview_Trigger\\oc_detils\\"
filepathxlsx = os.path.join(path, xlsx_file_name)    
filepathjson = os.path.join(path, json_file_name)

url = "http://api.fabtrakr.com/analytics/sewing/dateWise/efficiency/location/IDU3"
parameters = {
  "startDate" : date_input,
  "endDate" : date_input
}
payload={}
headers = {
  'version': '8.0',
  'Authorization': 'Bearer ' + str(token) 
  }


response = requests.request("GET", url,params=parameters, headers=headers, data=payload)
jsonData = json.loads(response.text)
with open(filepathjson, "w") as outfile:

    json.dump(jsonData, outfile)
p=Path(filepathjson)

with p.open('r', encoding='utf-8') as f:
   data = json.loads(f.read())
df = pd.read_json(filepathjson)

df.drop(['hourlyDetails'], axis=1, inplace = True)

writer = pd.ExcelWriter(filepathxlsx)
df.to_excel(writer)

writer.save()
