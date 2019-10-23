"""Calculates total volume of all walls in the model."""

#Set sys.path
import sys
sys.path.insert(0, 'C:\\Program files (x86)\\IronPython 2.7\\Lib')

# Python module
import traceback

# .Net module
import System

import ctypes

# Common Language Runtime module (clr)
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('RevitServices')
clr.AddReference('RevitNodes')
clr.AddReference('ProtoGeometry')

# StackGeneralNotes is what i need to "call" in the rjc.GeneralNotesAutomation "namespace". NEED TO CHANGE THE NAMESPACES IN CHRIS FEBRARRO'S GENERAL NOTES ADDIN TO NOT INCLUDE PERIODS!!

clr.AddReferenceToFileAndPath('C:\\Users\\acrossonbouwers\\Documents\\GitHub\\rjc-development\\rjcGeneralNotes2\\rjcGeneralNotesAutomation\\bin\\x64\\Debug\\GeneralNotesAutomation.dll')
from rjcGeneralNotesAutomation import StackGeneralNotes
myStackNotesInstance = StackGeneralNotes()

myStackNotesInstance.Start()



'''

from Autodesk.Revit import DB

doc = __revit__.ActiveUIDocument.Document

# Creating collector instance and collecting all the walls from the model
wall_collector = DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_Walls).WhereElementIsNotElementType()

# Iterate over wall and collect Volume data
total_volume = 0.0

for wall in wall_collector:
    vol_param = wall.Parameter[DB.BuiltInParameter.HOST_VOLUME_COMPUTED]
    if vol_param:
        total_volume = total_volume + vol_param.AsDouble()

# now that results are collected, print the total
print("Total Volume is: {}".format(total_volume))
'''