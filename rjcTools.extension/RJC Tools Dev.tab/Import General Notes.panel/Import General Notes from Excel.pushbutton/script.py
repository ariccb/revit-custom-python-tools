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

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document

user_excel_file = ''

def select_file():
    file_dialog = OpenFileDialog()
    file_dialog.InitialDirectory = "C:\Users\acrossonbouwers\Desktop\python test"
    file_dialog.ShowDialog()
    file_name = file_dialog.FileName    
    return file_name

user_excel_file = select_file()

print(user_excel_file)


def read_general_notes_form(): #start of the function that reads the General Notes Excel Form file...

    os.startfile(user_excel_file) #this opens Excel, the actual program (doesn't just read it from memory) and loads the selected file"""

    #Accessing the Excel applications.
    excel = Marshal.GetActiveObject("Excel.Application")



    #Opening a workbook
    workbook = excel.workbooks.Open(user_excel_file)    #excel files are called workbooks in the Marshal library

    #Worksheet, Row, and Column parameters
    worksheet = 1   #it's the first (and only) worksheet in the workbook
    rowStart = 2    #starts reading the excel file at row 2 to skip the titles for the columns
    rowEnd = 236    #this is how many rows total are in the RJC General Notes Selection Form.xlsm file - keep this updated if the form changes
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
    print list_generalnote_ids
    return list_generalnote_ids     # returning the list of general note id numbers

list_generalnote_ids = read_general_notes_form()    #assigning the list from read_general_notes_form() return value to a variable to be used later in the code


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
selectedViews = []
#similar list to spit out names for checking/clarity
selectedViews_Names = []


#lookup the parameter 'General Notes ID Number', using .AsString to get the value instead of just the Parameter, and add views to list if it matches any values in 'list_generalnote_ids'
for draftview in draftviews_collector: #creates a loop and iterates the following code through each instance in the list 'draftviews_collector' using draftview as the variable
    param = draftview.LookupParameter('General Notes ID Number').AsString()
    if param in list_generalnote_ids:
        selectedViews.append(draftview)         # saves the matching views to selectedViews list (the original element types)
        selectedViews_Names.append(draftview.Name)    # prints the view name

print selectedViews_Names



#commit the transaction to the Revit database
t.Commit()
 

