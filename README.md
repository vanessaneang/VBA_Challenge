# VBA_Challenge

## Overview of Project

The purpose of this project was to see if we could make our code run faster if we **_refactored_**, or edited, the code to loop through our stock analysis. In the end the result analyzed the stocks total daily volume and rate of return based on the stock tickers. 

## Steps and Code

### Create tickerIndex and Three Output Arrays

    1a) Create a ticker Index
 ```ruby   
    Dim tickerIndex As Integer
    tickerIndex = 0
 ```
    1b) Create three output arrays
 ```ruby      
    Dim TickerVolumes(12) As Long
    Dim TickerStartingPrices(12) As Single
    Dim TickerEndingPrices(12) As Single
 ```

### Create a *for* loop to initalize tickerVolumes= 0 then write a *for* loop to loop over all rows in spreadsheet, and increase current tickerVolumes to add current stock ticker volume

 
    2a) Create a for loop to initialize the tickerVolumes to zero.
 ```ruby 
    For i = 0 To 11
    
       TickerVolumes(i) = 0
       
    Next i
 
 ```
    2b) Loop over all the rows in the spreadsheet.
  
    For i = 2 To RowCount
    3a) Increase volume for current ticker
        TickerVolumes(tickerIndex) = TickerVolumes(tickerIndex) + Cells(i, 8).Value
     
    3b) Check if the current row is the first row with the selected tickerIndex.
     
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            TickerStartingPrices(tickerIndex) = Cells(i, 6).Value

     
    3c) check if the current row is the last row with the selected ticker
     
        ElseIf Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            TickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
  
    3d) Increase the tickerIndex.
   
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
        End If

### Use a for loop to loop through arrays to output the "Ticker","Total Daily Volume" and "Return" Columns 

    4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
```Ruby    
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = TickerVolumes(i)
        Cells(i + 4, 3).Value = TickerEndingPrices(i) / TickerStartingPrices(i) - 1
    
    Next i
```    
    
## Results

The original code only had one array defined for the tickers and one conditional for loop. It also formatted the table in a separate macro, so the time reflected for the 2017 and 2018 analyses does not reflect the entire time. 

![Original Code 2017 Run Time](https://github.com/vanessaneang/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017_notrefactored.png)

![Original Code 2018 Run Time](https://github.com/vanessaneang/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018_notrefactored.png)

Once the coded was refactored to include a TickerIndex it would organize that the three output arrays, TickerVolumes(), TickerStartingPrices(), and TickerEndingPrices() to output the correct values to the correpsonding ticker. Overall the output arrays did in fact increase the efficiency of the code, thus making the run time faster. Instead of placing the values individually into the cells, the arrays bypasssed this and could place the more variables at a time.

![Refactored Code 2017 Run Time](https://github.com/vanessaneang/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.png)

![Refactored Code 2018 Run Time](https://github.com/vanessaneang/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png)

There is a clear difference between the orginal code and the refactored code, by using the arrays to loop the multiple variables it runs about *_5 to 6 times faster_*.

## Summary

Refactoring code can lead to more efficient and concise code that is easier for both other coders to understand and for the computer to process. However, the disadvante to refactoring code would be the amount of time it takes to edit and increase efficiency for the code. It can also tamper with other parts of the code that works initially, complicating other parts of the code further. 

In the case of refactoring the code for VBA script it can lead to increase efficiency; the run time for the refactored code was 5-6 times faster than the orginal code. The process to edit and change the code did take some time, this would be the main disadvantage. In addition with VBA Script the code was not simipflied rather it was longer since more arrays and varaibles needed to be defined. Overall refactoring code can be advantageous in streamlining code, but the caveats may be more time spent trying to make a process slightly more effecient with more lines of code. 

