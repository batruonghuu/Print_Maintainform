import openpyxl
import tkinter.filedialog
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')

file_template_path = tkinter.filedialog.askopenfilename(title="Mở file để in")

wb = excel.Workbooks.Open(file_template_path)

exclude_sheet = 'Original Sheet'
# i = 0
for sheet in wb.Sheets:

      if sheet.Name!= exclude_sheet:
            # i += 1
      # if sheet.Name == 'KG-T1E-SCN-COREROOMT1E T1':
            print(sheet.Name)
            sheet.PrintOut()
            # print(i)