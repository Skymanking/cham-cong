import    win32com.client
Excel = win32com.client.Dispatch("Excel.Application")

# python -m pip install pywin32
file=r'baocao.xlsx'
wb = Excel.Workbooks.Open('baocao.xlsx')
sheet = wb.ActiveSheet

#Get value
val = sheet.Cells(1,1).value
# Get Formula
sheet.Cells(6,2).Formula