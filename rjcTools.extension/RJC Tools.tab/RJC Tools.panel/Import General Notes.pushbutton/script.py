#import libraries and reference the RevitAPI and RevitAPIUI
import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import * 

 
#set the active Revit application and document
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
 
#define a transaction variable and describe the transaction
t = Transaction(doc, 'This is my new transaction')
 
#start a transaction in the Revit database
t.Start()
 
#perform some action here...
print("I see you've tried to run the Import General Notes Tool.\nIt's still a work in progress, but it's going to be EPIC!\nSo please hang tight, it'll be released hopefully soon!")
#commit the transaction to the Revit database
t.Commit()
 

