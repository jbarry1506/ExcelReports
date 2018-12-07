import tkinter as tk
import openpyxl

"""
Create form in Tkinter: 
    open an excel file.
    Search for a user-defined word or phrase to use as a starting point
    copy all rows to the end of the document
    paste in new document

"""


wb = openpyxl.load_workbook(
        r'\\dsfiles01\Managed Services\GEMS Reporting\2018\rawFiles\SGS UCDay2_MSOpenandCompletedTickets.xlsx'
    )
sheetnames = wb.get_sheet_names()
ws = wb.get_sheet_by_name(sheetnames[0])

wsrows = ws.max_row
wscols = ws.max_column
print("wsrows " + str(wsrows))
print("wscols " + str(wscols))

def copycells(startrow, endrow, startcol, endcol, sheet):
    rangesel = []
    for i in range(startrow, endrow + 1, 1):
        rowsel = []
        for j in range(startcol, endcol + 1, 1):
            rowsel.append(ws.cell(row=i, column=j).value)
        rangesel.append(rowsel)

    return rangesel


def findtext(s_row, e_row, s_col, e_col, searchsheet, searchtext):
    cell_info = copycells(s_row, e_row, s_col, e_col, searchsheet)
    for c in cell_info:
        print(cell_info[c][1].value)

findtext(1, wsrows, 1, wscols, ws, "Ticket Number")


mainWindow = tk.Tk()

mainWindow.title("Did it work")
mainWindow.geometry("640x480")
mainWindow.mainloop()
