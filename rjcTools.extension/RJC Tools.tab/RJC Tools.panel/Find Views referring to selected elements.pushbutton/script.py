#Set sys.path
import sys
sys.path.insert(0, 'C:\\Program files (x86)\\IronPython 2.7\\Lib')

import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import * 

# Import Document Manager and TransactionManager
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

# Import RevitAPI
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *

from pyrevit.forms import WPFWindow

doc = DocumentManager.Instance.CurrentDBDocument
__title__='Find Views Referring\n To Selected Elements'

elementz = uidoc.Selection.PickObject(ObjectType.Element)

viewtemplates = list()
views = list()
viewcollector = FilteredElementCollector(doc).OfClass(View)

sheetcollector = FilteredElementCollector(doc).OfClass(ViewSheet)

for view in viewcollector:
    if view.IsTemplate == True:
        viewtemplates.append(view)
    else:
        views.append(view)

viewmatches = list()
for elem in elementz:
    viewmatchez = list()
    viewmatches.append(viewmatchez)

for v in views:
    viewelemidz=[]
    viewelems = list()
    viewelems = FilteredElementCollector(doc,v.Id)
    for viewelem in viewelems:
        viewelemidz.append(viewelem.ID)
    for i in range(0,len(elementz)):
        elem = elementz[i]
        elemid = elem.Id
        viewmatchez = viewmatches[i]
        if elemid in viewelemidz:
            viewmatchez.append(v)

sheetz = list()
for sheet in sheetcollector:
    sheetz.append(sheet)

elsheetmatches = list()
for i in range(0,len(elementz)):
    elviewmatches = viewmatches[i]
    elsheets = list()
    for viewmatch in elviewmatches:
        sheetnum = viewmatch.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER).AsString()
        if sheetnum != '---':
            for sheetel in sheetz:
                if sheetel.get_Parameter(BuiltInParameter.SHEET_NUMBER).AsString() == sheetnum:
                    elsheets.append(sheetel)
    elsheetmatches.append(elsheets)

OUT = viewmatches, elsheetmatches