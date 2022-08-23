# VBA_Challenge

## Overview of Project

The purpose of this project was to see if we could make our code run faster if we **_refactored_**, or edited, the code to loop through our stock analysis. In the end the result analyzed the stocks total daily volume and rate of return based on the stock tickers. 

## Results

The original code only had one array defined for the tickers and one conditional for loop. It also formatted the table in a separate macro, so the time reflected for the 2017 and 2018 analyses does not reflect the entire time. 

![Original Code 2017 Run Time](https://github.com/vanessaneang/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017_notrefactored.png)

![Original Code 2018 Run Time](https://github.com/vanessaneang/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018_notrefactored.png)

Once the coded was refactored to include a TickerIndex it would organize that the three output arrays, TickerVolumes(), TickerStartingPrices(), and TickerEndingPrices() to output the correct values to the correpsonding ticker. Overall the output arrays did in fact increase the efficiency of the code, thus making the run time faster. Instead of placing the values individually into the cells, the arrays bypasssed this and could place the more variables at a time.

![Refactored Code 2017 Run Time](https://github.com/vanessaneang/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.png)

![Refactored Code 2018 Run Time](https://github.com/vanessaneang/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png)

There is a clear difference between the orginal code and the refactored code, by using the arrays to loop the multiple variables it runs about 5 to 6 times faster.

The analysis is well described with screenshots and code (4 pt).

Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
Submission
