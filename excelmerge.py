import pandas as pd
import glob
from datetime import datetime
from datetime import date


date_today  = date.today() 
date_str = datetime.now().strftime('%Y-%m-%d')
date_input = datetime.now().strftime('%Y%m%d')
xlsx_file_combined =  'combined-'+date_str+'.xlsx'

# getting excel files to be merged from the Desktop 
path = "C:\\Users\\Gurumurthy\\Documents\\Overview_Trigger\\oc_detils\\"

# read all the files with extension .xlsx i.e. excel 
filenames = glob.glob(path + "\*.xlsx")

# empty data frame for the new output excel file with the merged excel files
outputxlsx = pd.DataFrame()

# for loop to iterate all excel files
for file in filenames:
   # using concat for excel files
   # after reading them with read_excel()
   df = pd.concat(pd.read_excel( file, sheet_name=None), ignore_index=True, sort=False)

   # appending data of excel files
   outputxlsx = outputxlsx.append( df, ignore_index=True)

outputxlsx.to_excel(xlsx_file_combined, index=False)