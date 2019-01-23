# -*- coding: utf-8 -*-
import glob
import string
import pandas as pd
import sys

reload(sys)
sys.setdefaultencoding('utf8')
count=1
excel_files=glob.glob('*.xlsx')
for ad in glob.glob('*.xls'):
    excel_files.append(ad)
for files in excel_files:
    try:
        xl = pd.ExcelFile(files)
        for sheets in xl.sheet_names:
            df= pd.read_excel(files, sheet_name=sheets)
            #print('Excel loaded to DataFrame')
            fields_check=[{'0':['UEN no','Company Registration Number','Organization UEN','Organizations UEN','Org UEN','Reg_No','REGISTRATION No','Registration Number','Registration_Number','UEN','UEN No','UEN No.','Uen No:','Uen Number','UEN_NO','UENNo','UENNumber','Company Registration Number']},
                          {'1':['Account Name','ACCOUNTS NAME','Company ','Company','Company Name','Company Name ','Company name as per ACRA','COMPANY- SG','Company_Name','EntityName','Hotel Name','Hotel Name as per ACRA','Name of Organisation','Name of the Company','Organiation Name','Organisation','Organisation Name','Organisation Name ( as per ACRA )','Organisation Name as per ACRA','Organization Name','Organization Name.1','Organization_Name','Organsiation Name','Partner Name / Account Name','Contacts Organization Name','Name_Of_Clinic']},
                          {'2':['Person 1 saluation','Saluation','SALUTAION','Salutaiton','Salutation','Salutation ','Salutation.1' ]},
                          {'3':['First Name','FIRSTNAME']},
                          {'4':['Last Name','LAST NAMEE','LASTNAME','Surname']},
                          {'5':['Contact Name','Contact Person','Contact Person ','Contact Person- 1','CONTACT PERSON-1','Contact_Name','Contact_Name_Full','Contacts Full Name','Full  Name','Full Name','Full Name.1','Key_Person_name_1','NAME OF THE PERSON','NEW CONTACT -1','Partner name','PERSON','Person Name','Person\'s Name','Primary Contact']},
                          {'6':['JOB TITLE','CONTACT TITLE','Contact_Designation ','Contacts Designation','Current Job Title','Desgination','Designation','Designation ','Designation / Job Title','Designation / Job Title.1','Designation/ Job Title','Designation/JobTitle','Job Functions','Job Tile / Designation','Job Tiles / Designation','Job Title','Key_Person_Designation_1','Person1 Designation','TITLE','TITLE / DESIGNATION','title.1','TITLES','POSITION']},
                          {'7':['Email id','Company Email','Contact_Email','Contacts Email ID','corrected email id','Email','Email Address','Email ID','Email ID Generated','Email ID.1','Email IDs','email.1','Email_Address','email_box','EMAIL_ID','EMAILID','General Mail ID','GENERIC EMAIL','Given_EMAIL ID','Mail ID','MAILING ADDRESS','New email id','OFFICIAL EMAIL','Official email.1','Person  Email','Person Email','Person1  Email']},
                          {'8':['Cell Phone','Company Phone','Company Phone.1','Direct Phone','Mobile','Person  Phone','Person 1 Phone','Person Phone','PH_NO','Phione ','Phone','Phone Niumber','Phone No','Phone Number','Phone Number-01','Phone_No','Phone_No.','Phone_Number','Phone001','Phone-001','PhoneNumber','Primary Phone','Tel','Tel_1','Telephone','Work Phone']},
                          {'9':['Company Fax','Facsimile','Facsmile','Fax','Fax ','Fax No','Fax Number','Fax_No','Fax_No.','Fax_Number','FaxNumber','TeleFax']},
                          {'10':['Company Website','Compay Website','WEB ADDRESS','WEB_ADDRESS','Website','Website ','Website / ( Client given list )','Website']},
                          {'11':['Company Email Domain','Domain']},
                          {'12':['add1','Address','ADDRESS ','ADDRESS #01','Address _ 01','ADDRESS _01','ADDRESS 01','Address 1','ADDRESS#01','Address_01','Address_1','Address-01','ADDRESS1','ADDRESS-1','Business_Address','CURRENT ADDRESS_01','Full  Address','Full address','Office_Address','STREET ADDRESS','SRegistered Address','Road Name','streetName']},
                          {'13':['Unit No','Unit Number','unitNumner']},
                          {'14':['Building','Building Name','buildingname','add2','ADDRES_02','ADDRESS #02','ADDRESS 02','Address 2','Address of Place of Practice','ADDRESS#02','Address_02','Address-02','ADDRESS2','ADDRESS-2','Address_2','CURRENT ADDRESS_02']},
                          {'15':['add3','Address_03','Address_3','City','Contact_Person Location','Location','Location of HR Person','Location of the HR person','STREET ADDRESS CITY']},
                          {'16':['Pin code','Post Code','POSTAL','Postal  Code','Postal Cde','Postal Code','Postal Number','Postal_Code','postalcode','POSTEL CODE','STREET ADDRESS ZIP','ZIP CODE','ZIPCODE']},
                          {'17':['Country','Country ','COUNTY','SCountry','Work Country']},
                          {'18':['1-Activity-Code-5','SSIC Description','Catagory','Category','industry','Industry ','INDUSTRY CATEGORY','Industry Description','Industry Description.1','Industry_001','Sector','Nature_Of_Business']},
                          {'19':['~ EMP Size','ACTUAL EMPLOYEE SIZE','Emp Size','EMP_SIZE','Employee size','EMPLOYEE SIZE RANGE','Employees','EMPsize','EMP-SIZE','Number_Of_Employees','OFFICE SIZE']},
                          {'20':['~ Sales Turnvoer/Capital (in $M)','~ SALES/REVENUE           ( in US Dollars) Million','REVENUE','Sales ( in Million USD)','Sales ( in USD)','Sales (in USD)','Sales Revenue','Sales Turnover','Sales Turnover Range(in Mil)','SALES VOLUME RANGE','Sales_Turnover','sales_TurnOver_In_Last_Fiscal_Year','ACTUAL SALES VOLUME','Net Worth (In $ Billion)']}
                          ]
            fields=['funknown0','funknown1','funknown2','funknown3','funknown4','funknown5','funknown6','funknown7','funknown8','funknown9','funknown10','funknown11','funknown12','funknown13','funknown14','funknown15','funknown16','funknown17','funknown18','funknown19','funknown20']
            for txts in fields_check:
                for txt in txts:
                    for t in txts[txt]:
                        #print(t)
                        for list_items in list(df):
                            if str(list_items).lower()==str(t).lower():
                                fields[int(txt)]=list_items
            #print('Index updated')
            #for cols in fields:
                #print(cols)
            for cols in fields:
                if cols.startswith('funknown'):
                    df[cols]=''
                    #df.insert(fields.index(cols),cols,'')
            df[fields].to_csv(string.replace(string.replace(files+'_'+sheets, '.xlsx', ''),'.xls','')+'.csv',sep=';',encoding='utf-8',index=None)
            print(str(count)+'  Completed File : '+string.replace(string.replace(files+'_'+sheets, '.xls', ''),'.xlsx','')+'.csv')
            count+=1
    except Exception as e:
        print('Error Accured.......')
        print(e)
