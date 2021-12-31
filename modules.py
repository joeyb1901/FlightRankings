import openpyxl as xl
import os
from FlightMember import FlightMember

def processWorkbook(directory, wb, Filename):
    # Setup Output File
    directoryOut = directory + '\ProcessedData'
    OUTPUT_FILE = os.path.splitext(Filename)[0] + '_Processed' + os.path.splitext(Filename)[1]
    wbOut = xl.Workbook()  # Create output workbook to store info and rankings on
    fltSheet = wbOut.active
    fltSheet.title = "Flight Info"  # First sheet for flight info
    # Loop through all flight members and generate their analysis
    rawNames = wb['Sheet1']
    roster = []
    avgRanks = []
    for col in range(2, rawNames.max_column + 1):  # Loop through columns
        cell = rawNames.cell(1, col)
        member = FlightMember(rawNames, wbOut, cell.value)
        roster.append(cell.value)  # Adds name to roster
        avgRanks.append(FlightMember.getAvg(member))  # Adds corresponding avg rank to list
    avgRanks, roster = orderSimultaneously(avgRanks, roster)  # Order lists according to avg rank
    # Add Flight Info
    flightInfo(fltSheet, avgRanks, roster)
    # Save workbook
    try:
        wbOut.save(os.path.join(directoryOut, OUTPUT_FILE))
    except FileNotFoundError:
        os.mkdir(os.path.join(directoryOut))
        wbOut.save(os.path.join(directoryOut, OUTPUT_FILE))  # Save workbook to ProcessedData


def orderSimultaneously(list1, list2):  # Ordered according to list 1
    zipped_lists = zip(list1, list2)
    sorted_pairs = sorted(zipped_lists)

    tuples = zip(*sorted_pairs)
    list1, list2 = [list(tuple) for tuple in tuples]
    return list1, list2


def flightInfo(ws, ranks, roster):
    header = {
        'A1': "Name:",
        'B1': "Avg Rank:"
    }
    # Add items in template to xlsx
    for item in header.items():
        ws[item[0]] = item[1]
    # Add ranking list to xlsx
    for row in range(2, len(roster) + 2):
        ws.cell(column=1, row=row, value=roster[row-2])
        ws.cell(column=2, row=row, value=ranks[row-2])
