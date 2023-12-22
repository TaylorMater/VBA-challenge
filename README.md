# VBA-challenge
Module 2 of Data Analytics Bootcamp

Riley Taylor
December 2023

#####################################
IMPORTANT:

The scripts run best via developer buttons in the desired workbook. This is because, while I could verify that the code worked and generated the desired results when using buttons and debugging, I found that sometimes the console would give me errors depending on where my cursor was in the VBA developer window. I wrote several Sub procedures and even a function to make the code look cleaner, and I don't think it's great at handling that. For reference, the excel files have buttons on their first worksheet with the macros assigned. 

If you choose to not use the buttons provided in the files provided, then PLEASE NOTE that the SCRIPTS ARE SAVED IN MODULE 1 OF THE alphabetical_testing.xlsm which lies in the Start_Code\Resources directory of this repo. 

All buttons reference this workbook.

I have also provided the script separately as a text file - Taylor_Riley_Scripts.txt - feel free to copy this in to test it that way if it doesn't work. 


#####################################
Script Overview
#####################################

# WorkbookMarketAnalysis 
Type: Sub
Arguments: 
Description: This is the ultimate procedure, that applies WorksheetMarketAnalysis to all worksheets in the current active Workbook, and generates results on each worksheet

# WorksheetMarketAnalysis
Type: Sub
Arguments:
Description: This is the heaviest procedure, and it includes the iteration through all valid rows in a formatted worksheet like the ones provided. It adds data to the current ticker as it progresses through the iteration. (ideally, we would use a Ticker object/struct for this, but that's beyond the scope of this course). It relies on a call to CheckNextRow() to determine if the next row contains a new ticker, and relies on StoreTicker() to store relevant data for a particular ticker into the generated columns. Finally, it relies on DeriveExtremes() to produce the desired calculated results. 

# CheckNextRow
Type: Function
Arguments: currentRow As Long, currentTicker As String
Description: This function checks the next row in an iteration to see if it has the same ticker (case = 1), a different ticker (case = 2), or is empty (case = 3) and returns the case number.

# StoreTicker
Type: Sub
Arguments: currentRow As Long, currentTicker As String, firstOpen As Double, lastClose As Double, runningVolume As LongLong, summaryRow As Long
Description: This procedure is essentially fed all of the fields pertaining to a "Ticker" object if one existed, and then calculates/formats that data and stores it in the sheet. This is run after we have concluded with all the rows corresponding to a specific ticker. Total Volume is formatted to not be scientific, and percentage is formatted as a 0.00%. 

# ResetWorksheetAnalysis
Type: Sub
Arguments:
Description: Used to reset the cells entered from StoreTicker and DeriveExtremes, including color formatting but not percentage formatting. 

# DeriveExtremes
Type: Sub
Arguments:
Description: Used to search for the greatest incrase, decrease, and highest volume. This could have been implemented via current Max/Min/Total variables in the WorksheetMarketAnalysis, but I didn't want to clog it up too much. A future refactor might make this uneeded. 

# ResetWorkbookAnalysis
Type: Sub
Arguments:
Description: Same as ResetWorksheetAnalysis, but for all worksheets in the workbook


#####################################
Source List:
#####################################

# Info on the worksheet object in VBA:
https://support.microsoft.com/en-au/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0

I used code here to help create the Sub procedure WorkbookMarketAnalysis


# on using percentage format in VBA:
https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage

I used the Cells().NumberFormat="0.00%" multiple times


# on using Max in VBA
https://stackoverflow.com/questions/31906571/excel-vba-find-maximum-value-in-range-on-specific-sheet

I used this in my DeriveExtremes - I didn't know you had to prefix it with Application.WorksheetFunction


# Info on using XLOOKUP in VBA
https://support.microsoft.com/en-us/office/xlookup-function-b7fd680e-6d10-43e6-84f9-88eae8bf5929

I used this in my DeriveExtremes


# Info on number formatting 
https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba

Deabted using different options for Total Volume in StoreTicker, ended up leaving as general