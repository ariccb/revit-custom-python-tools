#import libraries and reference the RevitAPI and RevitAPIUI
import clr
import System

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
excel = Excel.ApplicationClass()

clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import * 

 
#set the active Revit application and document
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
 
#define a transaction variable and describe the transaction
t = Transaction(doc, 'Import General Notes')
 
#start a transaction in the Revit database
t.Start()
 
#perform some action here...
print("I see you've tried to run the Import General Notes Tool.\nIt's still a work in progress, but it's going to be EPIC!\nSo please hang tight, it'll be released hopefully soon!")

#commit the transaction to the Revit database
t.Commit()
 

