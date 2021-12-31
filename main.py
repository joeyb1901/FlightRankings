import openpyxl as xl
import os
from modules import processWorkbook

directory = r'C:\Users\joeyb\iCloudDrive\S6_TTU Spring 2022\AFROTC\_FTP FLT Files'

# TODO: Save data from previous weeks to track progression throughout the semester

def main():
    # Loop through files in RawRankingData to process
    directoryIn = directory + '\RawRankingData'
    for filename in os.listdir(directoryIn):
        if filename.endswith(".xlsx"):
            INPUT_FILE = os.path.join(directoryIn, filename)
            wbIn = xl.load_workbook(INPUT_FILE)
            processWorkbook(directory, wbIn, filename)
        else:
            continue

    print("Processing complete.")

main()
