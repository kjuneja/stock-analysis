# Stock Analysis With Excel VBA

## Overview of Project

### Purpose
The purpose of this project was to refactor Excel VBA code to show the analysis of entire stock market over the last few years. We refactored the VBA code provided to us to loop through all the data one time in order to collect the same information provided. Then we determined whether refactoring our code successfully made the VBA script run faster. 

This project uses the Microsoft Excel tools and VBA Script to create technical analysis and writing skills to write a written report based on analysis.

This project will consits of technical analysis deliverable and a writen report to deliver the results.

* Deliverable 1: Refactor VBA code and measure performance. 

  * This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time
  
* Deliverable 2: A written analysis of the results (README.md)

## Results

### Analysis of Original Code and Refactored Code
I was provided with original code where I needed to refactor some code to check the efficienty of the data. 


Snippet of Code where refactored was done. 
```    
    '1a) Create a ticker Index
Dim tickerIndex As Integer
tickerIndex = 0

'1b) Create three output arrays
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single

''2a) Create a for loop to initialize the tickerVolumes to zero.
' If the next row’s ticker doesn’t match, increase the tickerIndex.
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
    
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

Next i

'4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i
```
After running the macro the following stock prices were shown for year 2017 and 2018.

_**Snippet of 2017 Stock Table.**_

![2017 Stocks](https://user-images.githubusercontent.com/25447945/124402425-ce3af280-dcf5-11eb-8534-f385d87039ba.png)

_**Snippet of 2018 Stock Table.**_

![2018 Stock ](https://user-images.githubusercontent.com/25447945/124402441-e6127680-dcf5-11eb-83a2-431f393e17ca.png)

The execution Times of refaactored code also as follows:

_**Snippet of 2017 Stock Execution TIme**_

<img width="425" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/25447945/124402515-4b666780-dcf6-11eb-87be-97c7fced65d6.png">

_**Snippet of 2018 Stock Execution TIme**_

<img width="421" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/25447945/124402550-8b2d4f00-dcf6-11eb-9e61-6e2c3a740722.png">

## Summary

- What are the advantages or disadvantages of refactoring code?
  
  Advantages of refactoring code is to help ua make our code cleaner and more organized. It also include design and software improvement, better faster way to   debug and issues that may arise. 
  
  Disadvantage of refactoring code can be that if the application is too long then it would be difficult to debug the issue and it will take a long time to figure out any issues with it.    

- How do these pros and cons apply to refactoring the original VBA script?
  
  The huge benefit of refactoring the VBA script for stock market data is that the application was running much master than the before one. The old one was taking about one secound to run, whereas the new one took like approximately 0.15 seconds to run. The pros for this one was that there were extra line of codes and having extra time of code increses the time for debugging any issues.  


