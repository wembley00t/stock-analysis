# Stock-Analysis with VBA

## Overview of Project
In this project, Steve is assisting his parents with analyzing returns from various green energy stocks in 2017 and 2018 to determine 
if his parents should diversify their portfolio from the original investment in DAQO represented by the ticker symbol "DQ."  

### Purpose
The purpose of this project is to use refactored code to determine if it provides more efficient results than the original code developed
with the first analysis.

The data used in this analysis will include the stock ticker symbol, total daily volume and the stock return based on the starting and ending
price from either 2017 or 2018.

## Results

### Updated Code

A sample of the code updated with this project is below.  The code updates allow the program to move more efficiently through the set up of the 
three arrays for output and the use of the arrays as the program loops through total volume, starting price and ending price for each ticker
symbol.

'''  

    '1a) Create a ticker Index
    
    tickerIndex = 0
    
    '1b) Create three output arrays for volume, starting price and ending price
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single 
  
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
         End If
                        
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
         End If           

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            
            End If            
          
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1       
        
    Next i 
    '''
   ### Output
   
   The output for 2017 and 2018 are shown below.
   
   ![VBA_Challenge_2017](https://user-images.githubusercontent.com/100876517/161454925-81b48425-99e4-4445-a644-13fbb6252e69.png)
   ![VBA_Challenge_2018](https://user-images.githubusercontent.com/100876517/161454933-29a802b6-dbec-4600-8cb6-f560c05e856e.png)
    
   
   The time to run the script for both 2017 and 2018 improved from the approximately .49 seconds run time in the original code to .10 to .12
   seconds with the refactored code.
   
   

## Summary


    Refactoring code has potential advantages and disadvantages.  Advantages of refactoring include improved code readability,
    reduced code complexity and potential improved performance.  Disadvantages of refactoring include introducing bugs or unintended
    consequences to code. Full testing of refactored code is critical.  The cost and time component of testing the refactored code
    should be compared against the potential benefits of the refactored code for a specific project.
    
    For this specific Stock Analysis project, there was no noted disadvantage to refactoring the code.  The refactored code improved
    performance through the use of the tickerIndex and creating the arrays for volume, starting price and ending price that used the
    tickerIndex in the FOR loop.  This performance improvement will be more meaningful if the file becomes larger due to additional
    stock data being added.
