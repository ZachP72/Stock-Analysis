# Stock-Analysis
Analyzing DQ Stock and other various renewable energy stock 

## Overview of Project

### Purpose
The purpose of this task was to create a Microsoft Excel VBA code and use it to analyze stock from 2017 and 2018 to determine whether they should be invested in. Afterwards, I was tasked to make the code run more efficently and speed up the calculations. 

### Background
The data given to me provided information on 12 different stocks. The information contained ticker value, that date the stock was issued, the opening and closing price, the highest and losest price, and the volume of stock. Our goal was to make sure we compiled that up, and created macros that would instantly give us the ticker, daily volume, and return percentages on each stock. 

## Results

### Analysis
Prior to refactoring any code, I created a rough draft on Notepad ++. This rough draft featured code that would create the input box, chart headers, create loops, and activate on the correct worksheet with accurate formatting.Here is what was created: 
>     '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
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
    
## Summary 

# Thoughts on Refactoring Code
Refactoring code is a double edge sword. On the positive side of things you've got the ability to make code cleaner and easier to understand. Not to mention the speed increase to the performance the code. By creating a more efficent piece of code you open up more opportunities such as software improvement, added functionality, and easier debugging instances if needed. That positive side of the coin comes at a cost. This cost being human error. By starting to try to slim code down you open up the possibilities of making mistakes and throwing things out of wack. In which case if you make that mistake then you've got to go through the debugging process and start all over which wastes time. 

# Thoughts on Refactoring VBA Script
Similarly to what I have posted above there are a variety of pros and cons. The diffence between code and VBA script is that I had the pleasure to refactor VBA Script. The upside was that the macro happened in 0.2 seconds in comparison to the full second it took prior to the refactoring. While this is faster and less clunky showing up onto the excel sheet, it had taken me much longer than .8 seconds to refactor the entire thing. As long is it would have provided the same information and made sense in the coding process, then I don't know how worth the time invested would be. Here are my two successful images of the macro speed. 
![This is an image](https://gyazo.com/a6010c093f18b33dc6421361226d1340)
![This is an image](https://gyazo.com/067fe32225f50cf21c6421367b160d09)
