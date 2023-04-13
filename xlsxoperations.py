import pandas as pd
import numpy as np
from openpyxl import load_workbook
from datetime import datetime

date_str = datetime.now().strftime('%Y-%m-%d')
xlsx_file_name =  'file-'+date_str+'.xlsx'

FilePath = (xlsx_file_name)
ExcelWorkbook = load_workbook(FilePath)
writer = pd.ExcelWriter(FilePath, engine = 'openpyxl')
writer.book = ExcelWorkbook

DataSample1= [['IDD1'], ['IDU1'],['IDU3'],['IDU4'],['IDU5'],['IDU6'],['IDU8'],['IDU9P1'],['NJU1P1'],['NJU1P2'],['NJU2'],['NJU3']]

SimpleDataFrame1=pd.DataFrame(data=DataSample1, columns=['location'])

print(SimpleDataFrame1)

SimpleDataFrame1.to_excel(writer, sheet_name = 'Summary')
writer.save()
