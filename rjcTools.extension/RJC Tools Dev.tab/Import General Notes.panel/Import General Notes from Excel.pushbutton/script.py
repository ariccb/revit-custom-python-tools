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
rowStart = 2
rowEnd = 236
checkboxColumn = 6

#making list to store General Note ID #'s
listGNIDs = [] 
#variable for GNID's
gNoteNum = ''

#This is where i need to check the 6th (checkbox) Column in the "General Notes Selection Form" if it's TRUE or FALSE. 
#If it is TRUE, then collect the GN ID#, and ADD IT TO listGNIDs
#This needs to run row by row, and append the GN ID# to the list

#Using a loop to read a range of values and print them to the console.
for i in range(rowStart, rowEnd): 

    #Worksheet object specifying the cell location to read.
    data = workbook.Worksheets(worksheet).Cells(i, checkboxColumn).Text
    if(data == "TRUE"):
        listGNIDs.append(workbook.Worksheets(worksheet).Cells(i, 2).Text)
        if(data == ''):
            continue    
        print data
print listGNIDs
           

t = Transaction(doc, 'Read Excel spreadsheet.')
t.Start()
# Write code here that interacts with Revit

#using this video to figure this out: https://www.youtube.com/watch?v=WU_D2qNnuGg&index=7&list=PLc_1PNcpnV5742XyF8z7xyL9OF8XJNYnv


#this is where you're going to put your ElementParameterFilter 
#FilterStringRule() --> we need 'FilterableValueProvidor(valueProvidor type)' + 'FilterStringRuleEvaluator (evaluator type)' + 'string (ruleString type)' + 'bool (caseSensitive type)'
#ElementParameterFilter(FilterStringRule

#view_param_id = ElementId(__getitem__('General Notes ID Number'))
#print (view_param_id)

#collect all drafting views in project
draftviews_collector = FilteredElementCollector(doc).OfClass(ViewDrafting).ToElements()
print draftviews_collector


#create list to add views from draftviews_collector that have a matching parameter value
viewstouse = []

#lookup the parameter 'Sheet Number' and add views to list if it matches 'NM-0201'
for view in draftviews_collector:
    param = view.LookupParameter('General Notes ID Number').AsString()
    print param
    if param == '0000-001':
        viewstouse.append(view)

print viewstouse


#get element Id's for these views



#commit the transaction to the Revit database
t.Commit()
 

