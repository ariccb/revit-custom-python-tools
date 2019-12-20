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
from System.Drawing import*


inputFile = ""

fileDialog = OpenFileDialog()
fileDialog.InitialDirectory = "C:\Users\acrossonbouwers\Desktop\python test"
fileDialog.ShowDialog()
inputFile = fileDialog.FileName
        
print(inputFile)


app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
 

#Accessing the Excel applications.
excel = Marshal.GetActiveObject("Excel.Application")

#Opening a workbook
workbook = excel.workbooks.Open(inputFile)

#Worksheet, Row, and Column parameters
worksheet = 1
rowStart = 1
rowEnd = 100
column = 2
 
#Using a loop to read a range of values and print them to the console.
for i in range(rowStart, rowEnd):
    #Worksheet object specifying the cell location to read.
    data = workbook.Worksheets(worksheet).Cells(i, column).Text
    print data



t = Transaction(doc, 'Read Excel spreadsheet.')
t.Start()
# Write code here that interacts with Revit

#commit the transaction to the Revit database
t.Commit()
 

