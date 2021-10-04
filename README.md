# VBA Stock Analysis

## Overview of Project
The purpose of this analysis is to determine stock performance over a given time 
period. The background of this project started from stock data showing returns in 2017
and 2018 respectively. The original code provided limited analysis, so refactoring the
code to provide endless analysis was required.

### Results
I created a for loop and set all tickerVolumes to zero then looped over all rows
in the spreadsheet.
          For i = 0 To 11
            tickerVolumes(i) = 0
            
            Next i
         For j = 2 To RowCount

Then by creating if statements, I was able to output a value for the volume based on
tickerIndex array. Also I gathered the tickerStartingPrices and tickerEndingPrices 
as the loop read through the rows to detect the next tickerIndex.
        
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
                   
   
         If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
            End If
       

         If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
            
         tickerIndex = tickerIndex + 1
        
        End If
    
    Next j

Using the original code, analysis took 1.8203 seconds to run for 2017 and 2.0859 
seconds to run for 2018.  The refactored code provided results for 2017 in 0.3437
seconds and for 2018 in 0.9453 seconds.

https://github.com/methelen/stock-analysis/blob/main/VBA_Challenge_2017.png

https://github.com/methelen/stock-analysis/blob/main/VBA_Challenge_2018.png

The results showed that 2017 was a better year for the limited data selection for 
these 12 indexes. The DQ ticker and SEDG stock index provided the highest return. 
In 2018, the ENPH and RUN were the only stock indexes with a positive return.

## Summary
The advantages of refactoring code is that the code is easier to understand and is
 less complex. The only disadvantage is that this may take more time overall. 
This was evident for me as it took me multiple hours to refactor the original VBA
code and work through multiple syntax and compile errors. Though the results of the 
analysis was unchanged after refactoring, I now better understand arrays and nested
for loops and how that simplifies writing code.
