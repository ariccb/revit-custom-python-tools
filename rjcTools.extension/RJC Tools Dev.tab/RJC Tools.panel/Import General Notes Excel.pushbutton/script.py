import clr
import System
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import * 
 
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
 
t = Transaction(doc, 'Read Excel spreadsheet.')
 
t.Start()
 
#Accessing the Excel applications.
xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject('Excel.Application')
 
#Worksheet, Row, and Column parameters
worksheet = 1
rowStart = 1
rowEnd = 100
column = 2
 
#Using a loop to read a range of values and print them to the console.
for i in range(rowStart, rowEnd):
    #Worksheet object specifying the cell location to read.
    data = xlApp.Worksheets(worksheet).Cells(i, column).Text
    print data

#commit the transaction to the Revit database
t.Commit()
 

