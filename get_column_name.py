import json
import glob
import string
import xlsxwriter
from xlutils.copy import copy
from xlrd import open_workbook
import pandas as pd

input_file_name = raw_input("Enter The Input file Name: ")
output_file_name =input_file_name+'.xls'
workbook = xlsxwriter.Workbook(output_file_name)
worksheet = workbook.add_worksheet()
workbook.close()
book_ro = open_workbook(output_file_name)
book = copy(book_ro)
sheet1 = book.get_sheet(0)
roww=0
coll=0
count=1
excel_files=glob.glob('*.xlsx')
for ad in glob.glob('*.xls'):
    excel_files.append(ad)
for files in excel_files:
    try:
        xl = pd.ExcelFile(files)
        for sheets in xl.sheet_names:
            xl_sheet= pd.read_excel(files, sheet_name=sheets)
            if xl_sheet.shape[0]>10:
                sheet1.write(roww,coll,files)
                sheet1.write(roww,coll+1,sheets)
                sheet1.write(roww,coll+2,xl_sheet.shape[0])
                coll+=2
                for column_names in list(xl_sheet):
                    sheet1.write(roww,coll+1,column_names)
                    coll+=1
                    book.save(output_file_name)
                roww+=1
                coll=0
            print(str(count)+' Completed File : '+files+'-'+sheets)
            count+=1
    except:
        pass
