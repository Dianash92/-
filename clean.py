#How I Used Pandas to Automate Cleaning Excel Files with Python
import pandas as pd
import pandas as pd
import datetime as datet
import numpy as np     # Python numerical
f=pd.read_excel('d:/python/data.xlsx')     #read excel file

# Read Data from Google Sheets into Pandas
sheetname='database'
sheetid='1EFXnl4usFjVtBSElfcWK-_UtUdjSALPljo6_vCFTNEk'
url = 'https://docs.google.com/spreadsheets/d/' + sheetid + '/gviz/tq?tqx=out:csv&sheet=' + sheetname
dbase =pd.read_csv(url)

# View  the top/bottom  five rows
# print(f.head())
# print(f.tail())

#print(f.shape)    #how many columns and rows you’re working with.

# clean names
for i in f.columns[range(0,3)]:     #columns [first_name,father_name,last_name]
    f[i]=f[i].str.replace('احمد','أحمد')
    f[i]=f[i].str.replace('[A-Za-z0-9]','') #replace alphanumric
    f[i] = f[i].str.replace('\d', '')       #replace digit decimal
    #f[i]=f[i].str.replace('[A-Za-z0-9.-@؟ ]', '')
    f[i]=f[i].str.replace('\W','')         #replace symbols
    f[i] = f[i].fillna('')               #Replace null values

f.insert(3,'full_name',f['first_name']+' '+f['father_name']+' '+f['last_name']) #insert coulmns(full name)
print(f['full_name'])

# clean date
f['date']=f['date'].dt.date# data type date
f['date']=np.where(f['date']>datet.date(2009,12,30),datet.date(2008,1,1),np.where(f['date']<datet.date(1925,1,1),datet.date(1960,1,1),f['date']))
print(f['date'])



# Save file to Excel to new location

f.to_excel('d:/python/diana.xlsx',index=False)