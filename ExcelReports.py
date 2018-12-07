import tkinter as tk
import openpyxl

"""
Create form in Tkinter: 
    open an excel file.
    Search for a user-defined word or phrase to use as a starting point
    copy all rows to the end of the document
    paste in new document

"""


print("Hello Jim")
wb = openpyxl.load_workbook(
        r'\\dsfiles01\Managed Services\GEMS Reporting\2018\rawFiles\SGS UCDay2_MSOpenandCompletedTickets.xlsx'
    )
sheetnames = wb.get_sheet_names()
print(sheetnames)
sheet = wb.get_sheet_by_name(sheetnames[0])
print(sheet)
rowcount = 0
for row in range(1, sheet.max_row+1):
    if sheet['A'+str(row)].value == 'Ticket Number':
        print("found it!")
        print(row)
        rowarray = []
        rowarray += row
        print(rowarray)
        rowcount += 1


mainWindow = tk.Tk()

mainWindow.title("Did it work")
mainWindow.geometry("640x480")
mainWindow.mainloop()
