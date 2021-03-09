#import libraries and reference the RevitAPI and RevitAPIUI
import clr
import math
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

clr.AddReferenceByPartialName('System.Windows.Forms')
import System.Windows
import Autodesk.Revit.DB as DB
 
#set the active Revit application and document
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selection = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()]
__title__= 'Revision Specific\nPrint Set'


 
#define a transaction variable and describe the transaction
t = Transaction(doc, 'Print Set by revision')
 
#start a transaction in the Revit database
t.Start()
 
# collecting the revisions in the active revit document, first printing the list of all the revisions, then printing how many revisions there are in total, then printing the element ID of the LAST revision. 
revisionId = DB.Revision.GetAllRevisionIds(doc)
print(revisionId)
revision_length = (revisionId.__len__())
print(revision_length)
print(revisionId[revision_length-1])

#this is where i left off - it's not working properly yet
revisionDesc = DB.Revision.Description
print(revisionDesc)


#figure out how show user a list of revisions to select from - try to find source code of PyRevit's 'revision to print set' functionality



#commit the transaction to the Revit database
t.Commit()
 
#close the script window
def alert(msg):
    TaskDialog.Show('PyRevit',msg)

def quit():
    __window__.Close()

