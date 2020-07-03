
import pandas as pd
from pandas import ExcelWriter
import datetime as dt
import csv
import xlsxwriter
import re


FILEPATH = "C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/SAP_Reports"
FILENAME = "V_UC_Report.txt"

# Output
DIR = 'C:/Users/grwillia/OneDrive - Adobe Systems Incorporated/Desktop/Python/Testing/SAP/excel_output/'
EXCEL = "V_UC_Report.xlsx"
EXCEL2 = "V_UC_Report2.xlsx"


'''CSV'''
with open(FILEPATH + '/' + FILENAME, newline='') as f:
    reader = csv.reader(f, delimiter='|')
    data = list(reader)

'''REGEX'''
dashes = re.compile(r'-+')
strings = re.compile(r'\w+')

#remove dashed lines
# for lst in data:
#     for item in lst:
#         for m in re.finditer(dashes, item):
#             lst.remove(item) 
#             #print(lst)

new_list = []
# for lst in data:
#     for item in lst:
#         for m in re.finditer(strings, item):
#             new_list.append(m)
#             #new_list = re.findall('(\\W+)', m)
#             #print(lst)

for list2 in data:
    for item in list2:
        #print(item)
        #print(item)
        #for m in re.finditer(strings, item):
        s = re.findall(strings, item)
        new_list.append(s)

print(new_list)

# #remove empty lists
# for next_item in data:
#     data = [x for x in data if x]
#     for i in next_item:
#         #next_item = [x for x in next_item if x]
#         next_item = filter(None, next_item)


#data2 = data[1:]

# removes dashed strings
# i=0
# for item in data:
#     # if '----------------------------------------------------------------------------------------------------------------------------------' in item:
#     #     item.remove('----------------------------------------------------------------------------------------------------------------------------------')
#     # if '---------------------------------------------------------------------------------------------------------------------------------' in item:
#     #     item.remove('---------------------------------------------------------------------------------------------------------------------------------')
#     # if '' in item:
#     #     item.remove('')
#     m = dashes.match(item[i])
#     if m in item[i]:
#         item.remove(m)
#     i += 1
#     print(item)

# data2 =  list(filter(lambda a: a != [], data)) # removes empty lists

# for item in data2:
#     print(item)




# with xlsxwriter.Workbook(DIR + EXCEL2) as workbook:
#     worksheet = workbook.add_worksheet()

#     for row_num, data2 in enumerate(data2):
#         worksheet.write_row(row_num, 0, data2)

'''PANDAS'''
# headers = ['Document',	'Doc Category', 'Item',	'Short Description', 'General', 'Delivery', 'Billing Doc', 'Price', 'Goods Mov']
# V_UC_DF = pd.DataFrame(columns=headers)

# V_UC_DF = pd.read_csv(FILEPATH + '/' + FILENAME, skiprows=4, sep='|', engine='python') # 

# cols = [c for c in V_UC_DF.columns if c.lower()[:7] != 'unnamed'] # drops the empty unnamed columns
# V_UC_DF = V_UC_DF[cols]
# #V_UC_DF.rename(columns={"Document   Doc.cat. ": "Data"}, errors="raise")

# # #V_UC_DF.drop([0, 0], inplace=True) # Drop first empty rows


# V_UC_DF.to_excel(DIR + EXCEL, index=False)
# print(V_UC_DF.head())