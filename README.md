# VBA_challenge
README FILE
Within this dataset you will be working to create a VBA script  to analyze and summarize stock ticker data.
The code loops over each worksheet in the Excel workbook and performs the following actions:
Initializes variables for volume, summary table row, yearly change, percent change, yearly open, yearly close, ticker, max percent increase, last row, row index, and start.
Sets the title and column headers for the worksheet.
Gets the total number of rows in the worksheet.
Loops over the data in the sheet, checking if the ticker symbol changes. If it does, the code performs the following actions:

a.	Gets the ticker symbol.

b.	Adds the final volume to the total.

c.	Calculates the yearly change and percent change.

d.	Applies color formatting to the Yearly Change column based on whether the change is positive or negative.

e.	Adds the yearly change, percent change, total volume, and ticker symbol to the summary table.

f.	Checks if the percent change is greater than the current maximum percent increase and updates the max percent increase and ticker symbol accordingly.

g.	Checks if the percent change is less than the current maximum percent decrease and updates the max percent decrease and ticker symbol accordingly.

h.	Checks if the total volume is greater than the current maximum total volume and updates the max total volume and ticker symbol accordingly.

Outputs the results of the analysis to the summary table.
