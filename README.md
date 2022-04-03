# Stock-Analysis with VBA
Stock Analysis Sample 


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
   ### Results 
    
    




  * In the "Kickstarter" worksheet, a new column labeled "Years" was added.
  * Using the Year() function, the year was extracted from the "Date Created Conversion Column."
  * A new pivot table based on the Kickstarter data was placed in the new worksheet, "Theater Outcomes by Launch Date."
  * The pivot table was filtered based on "Parent Category" and "Years."
  * The row labels were the months of January through December.
  * The column labels were "successful," "failed," and "canceled."
   ![image](https://user-images.githubusercontent.com/100876517/160254948-e34212d4-d1f6-4a9a-a13a-32c629c70e84.png)
  * The parent category was filtered for "theater."
  * The campaign outcomes were sorted in descending order.
  
  ![image](https://user-images.githubusercontent.com/100876517/160255018-946f93cc-41d5-4962-970b-746e0875c118.png)
  
A line chart was created from the pivot table labeled Theater Outcomes Based on Launch Date.
![image](https://user-images.githubusercontent.com/100876517/160254715-951c2d49-e2fa-4baa-be32-62d1c142c74c.png)

This chart was saved as a .png file and is part of the resources folder.







## Summary

- What are two conclusions you can draw about the Outcomes based on Launch Date?

  The most successful theater outcomes based on launch date was in the month of May closely followed by the month of June.  The number
  of canceled outcomes was low and fairly consistent over the 12 month period with a slight uptick reflected in the month of January.

- What can you conclude about the Outcomes based on Goals?
  
  Most of the activity or 85% occurs within the $0 to $9,999 goal range.  The most successful outcomes were those with a goal of $4,999 or less.
  This goal range of $0 to $4,999 reflected a 73% to 76% success rate.
   
- What are some limitations of this dataset?   
 
  This dataset does not reflect the reason some plays were successful and others were not even within the same goal range.  A subcategory for 
  type of play may give additional information.  

- What are some other possible tables and/or graphs that we could create?

 Other possible tables and/or graphs could include a further breakdown of the successful outcomes to show the number of backers and average donation
 compared to the failed outcomes.  A table or graph to see if there is a correlation between launch date and goal amount based on outcome 
 could also be helpful.  
