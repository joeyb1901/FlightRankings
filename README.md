This program is used to automate the processing of team ranking information in an environment where the team members (A.K.A "flight members") are tasked with ranking
their flight mates, including themself.

When run, the program processes the rankings found in the directoryIn folder (all .xlsx files) and creates a new folder called ProcessedData with a .xlsx for each week
of rankings as well as a summary that tracks each flight member's progression over the weeks.

INSTRUCTIONS:

1. Create a folder titled RawRankingData in a directory of your choice (you will have to copy this directory to the code).

2. Fill the folder with .xlsx files of ranking information according to the format described below. The files can be named anything, but will be processed
in alphabetical order (I recommend Week1.xlsx, Week2.xlsx, etc.).

3. Change the 'directory' variable in main.py to the directory that holds RawRankingData. Note that r' ' is not part of the directory but must be included. Put your
directory in between the quotations.

4. Run the program. It will print two INFO lines when the processing weekly .xlsx files are created and when the summary file is created. All generated files will be in 
a folder called ProcessedData in the same directory as RawRankingData. If the folder does not exist, it will be created. An ERROR line will be printed if duplicate names 
are in a ranking or if not all members were ranked by someone.

INPUT FILE FORMATTING:

On the active sheet of the .xlsx workbook (usually called 'Sheet1') input the rankings of each flight. Row 1 is the name of the member with their rankings below. Column 1 is
used to number the rows.

Example:

Members Williams  Johnson   Rodriguez
1       Johnson   Rodriguez Rodriguez
2       Williams  Williams  Johnson
3       Rodriguez Johnson   Williams


Let me know if anyone puts this to use and how it works out! Enjoy.
