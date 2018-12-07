import tkinter as tk
import openpyxl
from openpyxl import Workbook

"""
Create form in Tkinter: 
    open an excel file.
    Search for a user-defined word or phrase to use as a starting point
    copy all rows to the end of the document
    paste in new document

"""


# Open the original workbook and sheet to get data from
wb = openpyxl.load_workbook(
        r'\\dsfiles01\Managed Services\GEMS Reporting\2018\rawFiles\SGS UCDay2_MSOpenandCompletedTickets.xlsx'
    )
sheetnames = wb.get_sheet_names()
ws = wb.get_sheet_by_name(sheetnames[0])

wsrows = ws.max_row
wscols = ws.max_column

# Create a new sheet to paste data
testfilename = r'\\dsfiles01\Managed Services\GEMS Reporting\2018\rawFiles\EXCELTEST.xlsx'
twb = Workbook()
newsheetnames = twb.get_sheet_names()
tws = twb.get_sheet_by_name(newsheetnames[0])
tws.title = "Ticket info"


def copycells(startrow, endrow, startcol, endcol, sheet):
    rangesel = []
    for i in range(startrow, endrow + 1, 1):
        rowsel = []
        for j in range(startcol, endcol + 1, 1):
            rowsel.append(ws.cell(row=i, column=j).value)
        rangesel.append(rowsel)

    return rangesel


def pastecells(startrow, endrow, startcol, endcol, newsheet, data):
    rowcount = 0
    for i in range(startrow, endrow+1, 1):
        colcount = 0
        for j in range(startcol, endcol):
            newsheet.cell(row=i, column=j).value = data[rowcount][colcount]
            colcount += 1
        rowcount += 1


def findtext(s_row, e_row, s_col, e_col, searchsheet, searchtext):
    cell_info = copycells(s_row, e_row, s_col, e_col, searchsheet)
    row = 0
    for c in cell_info:
        row += 1
        for v in c:
            if v == searchtext:
                return row


def ticketdata():
    print("Processing....")
    tixlist = copycells(newstart, wsrows, 1, wscols, ws)
    print("tixlist.__len__ = " + str(tixlist.__len__()))
    print("tixlist[0].__len__ = " + str(tixlist[0].__len__()))
    #tws.max_row = tixlist.__len__()
    #tws.max_column = tixlist[0].__len__()
    pastecells(1, tixlist.__len__(), 1, (tixlist[0].__len__()+1), tws, tixlist)


    twb.save(testfilename)
    print("the file should be saved with data")


newstart = findtext(1, wsrows, 1, wscols, ws, "Ticket Number")
ticketdata()
mainWindow = tk.Tk()

mainWindow.title("Did it work")
mainWindow.geometry("640x480")
mainWindow.mainloop()
