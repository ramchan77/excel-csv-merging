import json
import glob
import string
import xlsxwriter
from xlutils.copy import copy
from xlrd import open_workbook

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
json_files=glob.glob('*.json')
for files in json_files:
    json_data=open(files)
    jdata = json.load(json_data)
    def get_keys(dl, keys_list):
        if isinstance(dl, dict):
            keys_list += dl.keys()
            map(lambda x: get_keys(x, keys_list), dl.values())
        elif isinstance(dl, list):
            map(lambda x: get_keys(x, keys_list), dl)
    keys = []
    get_keys(jdata, keys)
    key_names=list(set(keys))
    sheet1.write(roww,coll,files)
    print('Completed File : '+files)
    for field_names in key_names:
        #print(files+' : '+field_names)
        sheet1.write(roww,coll+1,field_names)
        coll+=1
        book.save(output_file_name)
    roww+=1
    coll=0

