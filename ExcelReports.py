import tkinter as tk
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Fill, Font, NamedStyle
from copy import copy, deepcopy


"""
Create form in Tkinter: 
    open an excel file.
    Search for a user-defined word or phrase to use as a starting point
    copy all rows to the end of the document
    paste in new document

"""

# Open the original workbook and sheet to get data from
# Create a function to iterate over all excel files in a directory
wb = openpyxl.load_workbook(
        r'\\dsfiles01\Managed Services\GEMS Reporting\2018\rawFiles\SGS UCDay2_MSOpenandCompletedTickets.xlsx'
    )
sheetNames = wb.get_sheet_names()
ws = wb.get_sheet_by_name(sheetNames[0])

wsRows = ws.max_row
wsCols = ws.max_column

# Create a new sheet to paste data
testFileName = r'\\dsfiles01\Managed Services\GEMS Reporting\2018\rawFiles\EXCELTEST.xlsx'
twb = Workbook()
newsheetnames = twb.get_sheet_names()
tws = twb.get_sheet_by_name(newsheetnames[0])
tws.title = "Ticket info"


# FORMATTING
bold_font = Font(bold=True)

def formatHeading(r, c):
    twsCell = tws.cell(r, c)
    twsCell.font = bold_font  # (font=twsCell.style.font(bold=True))

def formatTitle(r, c):
    pass

def formatSub():
    pass


def formatSubBold():
    pass


# dictionary of function options for ticket headings
ticketHeadings = {
    "Closed": formatHeading(1, 1)
}


ticketColList = []
typeColList = []
priorityColList = []
contactColList = []

ticketColSet = {}
typeColSet = {}
priorityColSet = {}
contactColSet = {}


def copycells(startrow, endrow, startcol, endcol, sheet):
    rangesel = []
    for i in range(startrow, endrow + 1, 1):
        rowsel = []
        for j in range(startcol, endcol + 1, 1):
            rowsel.append(ws.cell(row=i, column=j).value)

        rangesel.append(rowsel)
        # print(rowsel[0])  # some end rows have 'none' in the field

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
    tixlist = copycells(newstart, wsRows, 1, wsCols, ws)
    for t in tixlist:

        ticketColList.append(t[0])
        typeColList.append(t[2])
        priorityColList.append(t[3])
        contactColList.append(t[4])

        ticketColSet[t[0]] = tixlist.index(t)
        typeColSet[t[2]] = tixlist.index(t)

        if t[3] is not None:
            priorityColSet[t[3]] = tixlist.index(t)
        final_priority_set = sorted(priorityColSet)

        contactColSet[t[4]] = tixlist.index(t)

    print("tixlist.__len__ = " + str(tixlist.__len__()))
    print("tixlist[0].__len__ = " + str(tixlist[0].__len__()))

    # tws.max_row = tixlist.__len__()
    # tws.max_column = tixlist[0].__len__()
    rowNumber = 1
    columnNumber = 1

    # Company Name
    tws.cell(row=rowNumber, column=columnNumber).value = "Company"
    rowNumber += 1
    print("ws.cell(2, 1)" + str(ws.cell(2, 1).value))
    tws.cell(row=rowNumber, column=columnNumber).value = ws.cell(4, 1).value
    rowNumber +=2

    # Priority Ticket Breakdown
    tws.cell(row=rowNumber, column=columnNumber).value = "Total Tickets By Priority"
    rowNumber += 1
    tws.cell(row=rowNumber, column=columnNumber).value = "Priority"
    columnNumber += 1
    tws.cell(row=rowNumber, column=columnNumber).value = "Number of Tickets"
    rowNumber += 1
    columnNumber = 1
    uniqueTicketSet = set(ticketColSet)
    uniqueTypeSet = set(typeColSet)
    uniquePrioritySet = set(priorityColSet)
    uniqueContactSet = set(contactColSet)
    # for u in uniqueContactSet:
        # print("Unique: " + str(u))
    for p in final_priority_set:
        print(tws.cell(row=rowNumber, column=columnNumber).value)
        if tws.cell(row=rowNumber, column=columnNumber).value is not None:
            rowNumber += 1
            if str(p).find('1') != -1:
                print("There are " + str(priorityColList.count(p)) + " instances of ")
                print("Unique: " + str(p))
                tws.cell(row=rowNumber, column=columnNumber).value = str(p)
                columnNumber += 1
                tws.cell(row=rowNumber, column=columnNumber).value = str(priorityColList.count(p))
                columnNumber = 1
            elif str(p).find('2') != -1:
                print("There are " + str(priorityColList.count(p)) + " instances of ")
                print("Unique: " + str(p))
                tws.cell(row=rowNumber, column=columnNumber).value = str(p)
                columnNumber += 1
                tws.cell(row=rowNumber, column=columnNumber).value = str(priorityColList.count(p))
                columnNumber = 1
            elif str(p).find('3') != -1:
                print("There are " + str(priorityColList.count(p)) + " instances of ")
                print("Unique: " + str(p))
                tws.cell(row=rowNumber, column=columnNumber).value = str(p)
                columnNumber += 1
                tws.cell(row=rowNumber, column=columnNumber).value = str(priorityColList.count(p))
                columnNumber = 1
            elif str(p).find('4') != -1:
                print("There are " + str(priorityColList.count(p)) + " instances of ")
                print("Unique: " + str(p))
                tws.cell(row=rowNumber, column=columnNumber).value = str(p)
                columnNumber += 1
                tws.cell(row=rowNumber, column=columnNumber).value = str(priorityColList.count(p))
                columnNumber = 1
        else:
            if str(p).find('1') != -1:
                print("There are " + str(priorityColList.count(p)) + " instances of ")
                print("Unique: " + str(p))
                tws.cell(row=rowNumber, column=columnNumber).value = str(p)
                columnNumber += 1
                tws.cell(row=rowNumber, column=columnNumber).value = str(priorityColList.count(p))
                columnNumber = 1
            elif str(p).find('2') != -1:
                print("There are " + str(priorityColList.count(p)) + " instances of ")
                print("Unique: " + str(p))
                tws.cell(row=rowNumber, column=columnNumber).value = str(p)
                columnNumber += 1
                tws.cell(row=rowNumber, column=columnNumber).value = str(priorityColList.count(p))
                columnNumber = 1
            elif str(p).find('3') != -1:
                print("There are " + str(priorityColList.count(p)) + " instances of ")
                print("Unique: " + str(p))
                tws.cell(row=rowNumber, column=columnNumber).value = str(p)
                columnNumber += 1
                tws.cell(row=rowNumber, column=columnNumber).value = str(priorityColList.count(p))
                columnNumber = 1
            elif str(p).find('4') != -1:
                print("There are " + str(priorityColList.count(p)) + " instances of ")
                print("Unique: " + str(p))
                tws.cell(row=rowNumber, column=columnNumber).value = str(p)
                columnNumber += 1
                tws.cell(row=rowNumber, column=columnNumber).value = str(priorityColList.count(p))
                columnNumber = 1

    rowNumber += 2


    pastecells(rowNumber, tixlist.__len__(), 1, (tixlist[0].__len__()+1), tws, tixlist)

    twb.save(testFileName)
    print("the file should be saved with data")


newstart = findtext(1, wsRows, 1, wsCols, ws, "Closed")
ticketdata()

"""
mainWindow = tk.Tk()

mainWindow.title("Did it work")
mainWindow.geometry("640x480")
mainWindow.mainloop()
"""