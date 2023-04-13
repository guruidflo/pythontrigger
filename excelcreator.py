import requests
import json
import pandas as pd
import openpyxl
from pathlib import Path
from datetime import datetime,timedelta
from datetime import date
from openpyxl import Workbook
from firebasetokengenerator import token

date_today  = date.today() 
date_str = datetime.now().strftime('%Y-%m-%d')
date_input = datetime.now().strftime('%Y%m%d')
json_file_name = 'file-'+date_str+'.json'
xlsx_file_name =  'file-'+date_str+'.xlsx'


url = "https://api.fabtrakr.com/analytics/sewing/overView"
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
with open(json_file_name, "w") as outfile:
#print(jsonData)

    json.dump(jsonData, outfile)
p=Path(json_file_name)

with p.open('r', encoding='utf-8') as f:
   data = json.loads(f.read())

df = pd.read_json(json_file_name)
df.drop(['locationName','locationWorkDays','batchWorkDays'], axis=1, inplace = True)

writer = pd.ExcelWriter(xlsx_file_name)
df.to_excel(writer)


writer.save()

