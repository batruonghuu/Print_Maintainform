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

# df_groupttb['Thời gian dự kiến'] = df_groupttb['Thời gian dự kiến'].str.slice(stop=10)
# slice the string in datetime by keep determine character number

df_groupttb.to_excel('dataframe.xlsx', index=False)