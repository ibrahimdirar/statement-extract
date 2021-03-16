import csv
import tkinter as tk
from tkinter import filedialog
import os
import sys
from win32com.client import Dispatch

root = tk.Tk()
root.withdraw()

# list position for statements
account = 0
date = 1
type = 2
narrative = 3
payment = 5
receipt = 6
balance = 7


def processStatement(statement):
    # Get csv and return an array where each element is a line stored as a list
    results = []
    with open(statement) as csvfile:
        # change contents to floats
        reader = csv.reader(csvfile, dialect=csv.excel)
        saveToArray = False
        tmpArray = []
        olb = 0
        clb = 0
        for row in reader:  # each row is a list
            # Checks if the line is the end of an account
            if (row[narrative] == "Current Ledger Balance") or \
                (row[narrative] == "Closing Ledger Balance") or \
                    (row[narrative] == "Forecast Ledger Balance"):
                clb = row[balance]
                # Stop storing lines as we only care about transactions
                saveToArray = False
                # Only stores transactions if there is a change in balance
                if (olb != clb):
                    results.extend(tmpArray)
                # Resets the temporary array for the next account
                tmpArray = []
            if (saveToArray) and (row[type] != ""):
                tmpArray.append(row)
            if (row[narrative] == "Opening Ledger Balance"):
                saveToArray = True
                olb = row[balance]
    return results


# Stores the selected file into an array
combinedResults = []

combinedResults.extend(processStatement(filedialog.askopenfilename()))
combinedResults.extend(processStatement(filedialog.askopenfilename()))

# Table header
output = '"Date","Account","Type","Narrative","Payment","Receipt",\n'

# Writes data to table in the way we want then print to console
for row in combinedResults:
    output += row[date] + ',"' + row[account] + '","' + row[type] + \
        '","' + row[narrative] + '",' + \
        row[payment] + ',' + row[receipt] + '\n'


# print(output)

# Save to file
with open("output.csv", "w") as csvfile:
    print(output, file=csvfile)

# Saves to excel
excel = Dispatch('Excel.Application')
excel.Visible = True

# Opens the excel file, bundled version will have sys.frozen set to True
if getattr(sys, 'frozen', False):
    # running in a bundle
    # os.system('start excel.exe "%s\\output.csv"' % (sys._MEIPASS, ))
    wb = excel.Workbooks.Open("%s\\output.csv" % (os.path.abspath("."), ))
else:
    # running live
    wb = excel.Workbooks.Open("%s\\output.csv" % (sys.path[0], ))
    # os.system('start excel.exe "%s\\output.csv"' % (sys.path[0], ))

# Activate first sheet
excel.Worksheets(1).Activate()

# Autofit column in active sheet
excel.ActiveSheet.Columns.AutoFit()

# Save changes in a current file
wb.Save()
