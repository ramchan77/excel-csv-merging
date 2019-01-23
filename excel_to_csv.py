import json
import glob
import string
import xlsxwriter
from xlutils.copy import copy
from xlrd import open_workbook
import pandas as pd

input_file_name = raw_input("Enter The Input file Name: ")
output_file_name =input_file_name+'.csv'
count=1
excel_files=glob.glob('*.xlsx')
for ad in glob.glob('*.xls'):
    excel_files.append(ad)
xl_sheet= pd.DataFrame()
for files in excel_files:
    try:
        xl_sheet1= pd.read_excel(files,index_col=0,header=None)
        xl_sheet=pd.concat([xl_sheet,xl_sheet1])
        print(str(count)+' Completed File : '+files)
        count+=1
    except Exception as e:
        print(e)
xl_sheet.to_csv(output_file_name, encoding='utf-8')
