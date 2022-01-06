import openpyxl as xl
import os
from modules import processWorkbook, createSummary

directory = r'C:\Users\joeyb\iCloudDrive\S6_TTU Spring 2022\AFROTC\_FTP FLT Files'

# TODO: Account for cadets dropping midway through the semester
# TODO: Catch errors in input workbooks such as repeated or missing names

def main():
    # Loop through files in RawRankingData to process
    directoryIn = directory + '\RawRankingData'
    summaryRoster = []  # Will be list of rosters each week
    summaryRanks = []  # Will be list of avg ranks corresponding to weekly roster names (regardless of order)
    summarySelfRanks = []
    for filename in os.listdir(directoryIn):
        if filename.endswith(".xlsx"):
            INPUT_FILE = os.path.join(directoryIn, filename)
            wbIn = xl.load_workbook(INPUT_FILE)
            fltRoster, fltRanks, fltSelfRanks = processWorkbook(directory, wbIn, filename)  # Process the current wb
            summaryRoster.append(fltRoster)
            summaryRanks.append(fltRanks)
            summarySelfRanks.append(fltSelfRanks)
        else:
            continue

    print("[INFO] Individual processing complete.")

    createSummary(directory, summaryRoster, summaryRanks, summarySelfRanks)

    print("[INFO] Summary complete.")


main()
