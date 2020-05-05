import clr
import System
import os
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
clr.AddReference("Microsoft.Office.Interop.Excel")
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.DB import * 
from System.Windows.Forms import *
from System.Runtime.InteropServices import Marshal
from Microsoft.Office.Interop import Excel
from System.Drawing import*

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document


inputFile = ""
fileDialog = OpenFileDialog()
fileDialog.InitialDirectory = "C:\Users\acrossonbouwers\Desktop\python test"
fileDialog.ShowDialog()
inputFile = fileDialog.FileName
        
print(inputFile)


ex = Excel.ApplicationClass()   
ex.Visible = True
ex.DisplayAlerts = False   

#Worksheet, Row, and Column parameters
worksheet = 1
rowStart = 1
rowEnd = 100
column = 2

workbook = ex.Workbooks.Open(inputFile)
ws = workbook.Worksheets[1]

print ws.Rows[1].Value2[0,0]


'''
# use this structure for when doing things within revit inside a transaction
t = Transaction(doc, 'Read Excel spreadsheet.')
t.Start()
# Write code here that interacts with Revit

#commit the transaction to the Revit database
t.Commit()
'''

