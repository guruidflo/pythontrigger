import pandas as pd
from datetime import datetime
from datetime import timedelta
from datetime import date
import matplotlib.backends.backend_pdf
import matplotlib.pyplot as plt

date_str = datetime.now().strftime('%Y-%m-%d')
json_file_name = 'file-'+date_str+'.json'
xlsx_file_name =  'file-'+date_str+'.xlsx'
df = pd.read_excel(xlsx_file_name, sheet_name='Summary')

# Create a PDF file with matplotlib
pdf = matplotlib.backends.backend_pdf.PdfPages("Summary.pdf")

# Plot the DataFrame as a table
plt.figure(figsize=(8,6))
plt.axis('tight')
plt.axis('off')
the_table = plt.table(cellText=df.values,colLabels=df.columns,loc='center')
pdf.savefig()

# Close the PDF file
pdf.close()
