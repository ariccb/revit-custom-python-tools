# Set sys.path
import sys
sys.path.insert(0,'C:\Program Files (x86)\IronPython 2.7\Lib')

#Python module
import traceback

# .Net module
import System
from System import Environment
import os

# Common Language Runtime Module
import clr
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
clr.AddReference("Microsoft.Office.Interop.Excel") 
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")

from Autodesk.Revit.ApplicationServices import *
from Autodesk.Revit.DB import * 
from System.Windows.Forms import *
from Microsoft.Office.Interop import Excel
from Microsoft.Office.Interop.Excel import ApplicationClass
from System.Runtime.InteropServices import Marshal
from System.Drawing import*

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
__title__='Import Corporate General\nNotes From\nExcel'

user_excel_file_path = ''
doc_path = str(BasicFileInfo.Extract(doc.PathName).CentralPath)
corp_metric_gn_path = 'R:\Technical Resources\STR\General Notes and Details\Revit\STR-STD-001-20200812-RCK-Metric Notes - Revit 2018.rvt'
corp_imp_gn_path = 'R:\Technical Resources\STR\General Notes and Details\Revit\STR-STD-000-20200812-RCK-Imperial Notes - Revit 2018.rvt'
cal_gn_path = 'R:\Office Services\CAL\STR\Cal Project Resources\_Production Resources\3 DD\RJC Calgary General Notes\RJC Calgary General Notes - Metric.rvt'
print(doc_path)
                        ###OPTIONAL FEATURE ADD###
                        #could add "pick from list" within revit by using pyrevit.forms
                        #for reference, open __init__.py at C:\Users\acrossonbouwers\AppData\Roaming\pyRevit-Master\pyrevitlib\pyrevit\forms
                        #Maybe it could ask if you want to load a pre-saved form (from engineer) or to load a list to select from yourself

def select_file():
    dlgOpen = OpenFileDialog()   
    dlgOpen.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) ##THIS ISN'T WORKING, NOT STARTING DIALOG AT CENTRAL FILE LOCATION
    dlgOpen.ShowDialog()
    file_name = dlgOpen.FileName  
    return file_name

user_excel_file_path = select_file()

print(user_excel_file_path)

def read_excel_general_notes_form():
    #Accessing the Excel applications.
    try:
        excel = Marshal.GetActiveObject("Excel.Application")
    except:
        excel = None
    
    if (excel == None):
        excel = ApplicationClass()
    
    #Opening a workbook
    workbook = excel.workbooks.Open(user_excel_file_path)    #excel files are called workbooks in the Marshal library
    
    #Worksheet, Row, and Column parameters NEEDS TO BE UPDATED IF THE FORM GETS LONGER!
    worksheet = 1   #it's the first (and only) worksheet in the workbook
    rowStart = 2    #starts reading the excel file at row 2 to skip the titles for the columns
    rowEnd = 500    #this is how many rows total are in the RJC General Notes Selection Form.xlsm file - keep this updated if the form changes
    checkboxColumn = 6  #defines which column contains the true/false values (here it's column 'F', which is the 6th column)

    #initializing a list to store General Note ID #'s
    list_generalnote_ids = [] 

    for i in range(rowStart, rowEnd): #Using a loop to read a range of values and print them to the console.

        #Worksheet object specifying the cell location to read.
        data = workbook.Worksheets(worksheet).Cells(i, checkboxColumn).Text
        if(data == "TRUE"):
            list_generalnote_ids.append(workbook.Worksheets(worksheet).Cells(i, 2).Text)
        if(data == ''):
            continue            
    print ("The selected views in the form are: ", list_generalnote_ids) #print which ones were checked in the excel file 
    return list_generalnote_ids     # returning the list of general note id numbers

list_generalnote_ids = read_excel_general_notes_form()    #assigning the list from read_general_notes_form() return value to a variable to be used later in the code


t = Transaction(doc, 'Import General Notes')
t.Start()
# Write code here that interacts with Revit

                    #testing this part 
                    # using this video to figure this out: https://www.youtube.com/watch?v=WU_D2qNnuGg&index=7&list=PLc_1PNcpnV5742XyF8z7xyL9OF8XJNYnv

                    #this is where you're going to put your ElementParameterFilter 
                    #FilterStringRule() --> we need 'FilterableValueProvidor(valueProvidor type)' + 'FilterStringRuleEvaluator (evaluator type)' + 'string (ruleString type)' + 'bool (caseSensitive type)'
                    #ElementParameterFilter(FilterStringRule

#collect all drafting views in project
draftviews_collector = FilteredElementCollector(doc).OfClass(ViewDrafting).ToElements()

#uncomment the next two lines if you want to check the element id's for the views in the filter
#draftviews_ids = FilteredElementCollector(doc).OfClass(ViewDrafting).ToElementIds()
#print draftviews_ids

#create list to add views from draftviews_collector that have a matching parameter value for the following for loop
matchedViews = []
#similar list to spit out names for checking/clarity
matchedViews_Names = []

# input('What General Note View ID Should I try to load?')
#lookup the parameter 'RJC Standard View ID', using .AsString to get the value instead of just the Parameter, and add views to list if it matches any values in 'list_generalnote_ids'
# try: 
#     for draftview in draftviews_collector: #creates a loop and iterates the following code through each instance in the list 'draftviews_collector' using draftview as the variable
#         param = draftview.LookupParameter('RJC Standard View ID').AsString()
#         if param in list_generalnote_ids:
#             matchedViews.append(draftview)         # saves the matching views to matchedViews array (the original element types)
#             matchedViews_Names.append(draftview.Name)    # saves the view name to the matchedViews_Names array
#             print(matchedViews_Names)
#     print ("Views in project matching the selection form are: " + matchedViews_Names)   # prints the selctedViews_Names array with added view names from ^ for loop (this array was empty prior to the for loop)

for draftview in draftviews_collector: #creates a loop and iterates the following code through each instance in the list 'draftviews_collector' using draftview as the variable
    param = draftview.LookupParameter('RJC Standard View ID').AsString()
    if param in list_generalnote_ids:
        matchedViews.append(draftview)         # saves the matching views to matchedViews array (the original element types)
        matchedViews_Names.append(draftview.Name)    # saves the view name to the matchedViews_Names array
    
print ("Views in project matching the selection form are: ")   # prints the selctedViews_Names array with added view names from ^ for loop (this array was empty prior to the for loop)
print (matchedViews_Names)
# except:
#     print("Looks like your project doesn't have the 'RJC Standard View ID' parameter loaded. This add-in can't continue, so it's exiting.")




#commit the transaction to the Revit database
t.Commit()
 

# need to update this program to:
# - Ask User for Metric or Imperial Version to Import
# - Search the Corporate File for matched "RJC Standard View ID" instead of the currently open 'doc'
# - Copy Views from Corporate File using "RJC Standard View ID" Parameter
# - Create the Project Paramter if it doesn't exist in the project
