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

#DEFINE A CLASS  ----------------------------------------------------
class LoadExcelForm:
    
    #first method 'LoadFile'
    def LoadExcelFile(self, inputFile):
        #'fileDialog' property
        self.fileDialog = OpenFileDialog()
        fileDialog.InitialDirectory = "C:\Users\acrossonbouwers\Desktop\python test"
        fileDialog.ShowDialog()
        inputFile = fileDialog.FileName

        
               
        print(inputFile)

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
                


#MAIN CODE HERE -----------------------------------------------------
def main():
    #do stuff here 
 

    t = Transaction(doc, 'Read Excel spreadsheet.')
    t.Start()
    # Write code here that interacts with Revit

    #collect views using listGNIDs list to search for the corresponding parameter


    #commit the transaction to the Revit database
    t.Commit()
    
#the __name__ variable equals the __main__ by default
if __name__ == "__main__":
    #run main
    main()