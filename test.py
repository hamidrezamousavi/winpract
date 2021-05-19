from win32com.client import Dispatch
import os
path = os.getcwd()
filepath = path+'\\t.xlsx'


xlApp = Dispatch("Excel.Application")
xlApp.Visible = 1
xlApp.Workbooks.Open(filepath)
if xlApp.Workbooks('t1').Sheets('Sheet2').Cells(1,1).Value == 0:
    xlApp.Workbooks('t1').Sheets('Sheet1').Cells(1,1).Interior.Color = 150

sheet = xlApp.Workbooks('t1').Sheets(1)
sheet.Cells(2,2).Formula = '=A2*2'


xlApp.Workbooks(1).SaveAs(Filename = filepath)
xlApp.Workbooks(1).Close()

