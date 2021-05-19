from win32com.client import Dispatch
import os
path = os.getcwd()
filepath = path+'\\t.xlsx'


xlApp = Dispatch("Excel.Application")
xlApp.Visible = 1
xlApp.Workbooks.Add(filepath)
if xlApp.Workbooks('t1').Sheets('Sheet2').Cells(1,1).Value == 0:
    xlApp.Workbooks('t1').Sheets('Sheet1').Cells(1,1).Interior.Color = 255



