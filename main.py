import xlwings as xl
import os
import tkinter.filedialog
import xlrd
import openpyxl
import pandas as pd

file_template_path = tkinter.filedialog.askopenfilename(title="Mở file template phiếu bảo dưỡng mẫu")
# ask directory

file_groupttb_path = tkinter.filedialog.askopenfilename(title="Mở file group TTB")

df_groupttb = pd.read_excel(file_groupttb_path)
# create a pandas dataframe from file excel

df_groupttb.fillna(method='ffill', inplace=True)
# unmergce mergedcell to seperate cells with same value

df_groupttb[['Mã trang thiết bị','Tên trang thiết bị']] = df_groupttb['Hệ thống/Trang thiết bị'].str.split(":",expand=True)
# seperate 1 column "text-to-column" 2 column

df_groupttb['Mã trang thiết bị'], df_groupttb['Tên trang thiết bị'] = df_groupttb['Tên trang thiết bị'],df_groupttb['Mã trang thiết bị']
# swap the column 'Tên trang thiết bị' to position before 'Mã trang thiết bị'

df_groupttb.drop(df_groupttb[df_groupttb['Tên trang thiết bị'] == '»'].index, inplace = True)
# delete row if value in column include >>

df_groupttb = df_groupttb[df_groupttb['Mã trang thiết bị'].notnull()]
# delete row if value is null

df_groupttb['Thời gian dự kiến'] = pd.to_datetime(df_groupttb['Thời gian dự kiến'], errors='coerce',dayfirst=True)
df_groupttb['Thời gian dự kiến'] = df_groupttb['Thời gian dự kiến'].dt.strftime('%Y-%m-%d')
# convert datetime column
# first: convert to datetime format (keep non-datetime format cell),
# second: convert to strings,

set_datetime = set(df_groupttb['Thời gian dự kiến'].unique())
list_datetime = list(set_datetime)
print(list_datetime)

df_groupttb.to_excel('dataframe.xlsx', index=False)
# Save to a excel
ws = xl.Book(file_template_path).sheets.active
def insert_new_row(brow,sheet,value1,value2):
    sheet.range(brow,1).api.EntireRow.Insert()
    # insert a new row with keeping format of previous row
    sheet.range(brow,2).value = value1
    sheet.range(brow,3).value = value2
    sheet.range(brow,1).autofit()
    # autofit row based on content
xxx = 8
df_copy = df_groupttb[df_groupttb['Thời gian dự kiến'] == list_datetime[0]]
for check in df_copy['Mã trang thiết bị']:
    insert_new_row(xxx,ws,check,'yyy')
    xxx = xxx + 1



# insert_new_row(8,ws,'xxx','yyy')



