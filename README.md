# VBA-challenge

## Problem
We were asked to write a VBA script that analyzes multiple worksheets' worth of stock data. The number of worksheets and size of each worksheet was to be treated as a dynamic value. 

## Goal (copied from the assignment)
Create a script that loops through all the stocks for one year and outputs the following information:
* The ticker symbol.
* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
* The total stock volume of the stock.

## Setup
I decided it was best practice to separate out the steps of this script into two subroutines and four functions. One subroutine to iterate through the workbook's worksheets (as well as create/delete-recreate a summary sheet) and the other subroutine to perform the analysis for a given worksheet. Three of the functions was used to calculate maximum and minimum values when given a range and the last function was for determing if a worksheet already existed.

The analysis subroutine uses 'refRow' and 'refColumn' integer variables to dictate the analysis results positioning. The iterator used to find stock ticker transitions uses a three pronged approach: it is capable of detecting if it is at the start of, middle of, or end of a section. The data set is convenient as the ticker symbols are all grouped together. If they had not been, I would have used additional arrays to store and collate the information. 

## Execution Notes
The script's execution time is considerable. Even when using Application.ScreenUpdating it takes a minute or so. But the script runs reliably well as long as the source worksheets maintain their current layout.

## Limitations
The script will work as along as there are column headers and the 'refRow' value is not set below 2. If there is another worksheet named "Summary" in a workbook when this script is run, that worksheet will be deleted. 