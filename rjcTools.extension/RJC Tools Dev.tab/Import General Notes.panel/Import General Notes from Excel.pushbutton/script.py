# Set sys.path
import sys
sys.path.insert(0,'C:\Program Files (x86)\IronPython 2.7\Lib')

#Python module
import traceback

# .Net module
import System
import os

# Common Language Runtime Module
import clr
clr.AddReference('RevitAPI') 
clr.AddReference('RevitAPIUI') 
clr.AddReference("Microsoft.Office.Interop.Excel") 
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")



from Autodesk.Revit.DB import * 
from System.Windows.Forms import *
from Microsoft.Office.Interop import Excel
from Microsoft.Office.Interop.Excel import ApplicationClass
from System.Runtime.InteropServices import Marshal
from System.Drawing import*

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

user_excel_file_path = ''

def select_file():
    file_dialog = OpenFileDialog()
    file_dialog.InitialDirectory = "C:\Users\acrossonbouwers\Desktop\python test"
    file_dialog.ShowDialog()
    file_name = file_dialog.FileName    
    return file_name

user_excel_file_path = select_file()

print(user_excel_file_path)


def read_excel_general_notes_form(): #start of the function that reads the General Notes Excel Form file...

    #os.startfile(user_excel_file_path)   #this opens Excel, the actual program (doesn't just read it from memory) and loads the selected file
                                    #the Marshal.GetActiveObject function needs excel to actually be running to work - there's probably a way
                                    #to get this working in the background too... i just don't know it

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
    #variable for GNID's
    gen_note_numbers = ''

    #This is where i need to check the 6th (checkbox) Column in the "General Notes Selection Form" if it's TRUE or FALSE. 
    #If it is TRUE, then collect the GN ID#, and ADD IT TO list_generalnote_ids
    #This needs to run row by row, and append the GN ID# to the list

    #Using a loop to read a range of values and print them to the console.
    for i in range(rowStart, rowEnd): 

        #Worksheet object specifying the cell location to read.
        data = workbook.Worksheets(worksheet).Cells(i, checkboxColumn).Text
        if(data == "TRUE"):
            list_generalnote_ids.append(workbook.Worksheets(worksheet).Cells(i, 2).Text)
            if(data == ''):
                continue    
            #print data
    print "The selected views in the form are: ", list_generalnote_ids
    return list_generalnote_ids     # returning the list of general note id numbers

list_generalnote_ids = read_excel_general_notes_form()    #assigning the list from read_general_notes_form() return value to a variable to be used later in the code


t = Transaction(doc, 'Read Excel spreadsheet.')
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


#lookup the parameter 'General Notes ID Number', using .AsString to get the value instead of just the Parameter, and add views to list if it matches any values in 'list_generalnote_ids'
for draftview in draftviews_collector: #creates a loop and iterates the following code through each instance in the list 'draftviews_collector' using draftview as the variable
    param = draftview.LookupParameter('RJC View ID').AsString()
    if param in list_generalnote_ids:
        matchedViews.append(draftview)         # saves the matching views to matchedViews array (the original element types)
        matchedViews_Names.append(draftview.Name)    # saves the view name to the matchedViews_Names array

print "Views in project matching the selection form are: ", matchedViews_Names   # prints the seelctedViews_Names array with added view names from ^ for loop (this array was empty prior to the for loop)



#commit the transaction to the Revit database
t.Commit()
 

