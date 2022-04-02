# VBA Challenge

## Overview of Project

### Purpose
Throughout the module, I created macros to allow Steve to analyze two sets of stock information at the click of a button. In order to allow him to expand his data sets and continue using the workbook in the future, the VBA code needed to be refactored to handle larger amounts of data and reduce the amount of time the analysis takes to complete.

## Results

### Refactoring the Code
The largest change made to the original code was to create three ouput arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to store the data for when the loops run. I also kept the array from the original code (tickers) that assigned a numerical value to each of the 12 stock ticker options so they can be easily inserted into the loop code. To match the three new arrays to the tickers array, I created a variable named tickerIndex. From here, I was able to use the majority of the original code to create the nested loops. I also added a line to the refactored code to automate the increase of the tickerIndex value at the end of the second (nested) For loop. Examples of this section of both codes are below for comparison:

#### Original
```
'Prepare for analysis of tickers
        'initialize variables for starting and ending price
        Dim startingPrice As Single
        Dim endingPrice As Single
        
        'activate data WS
        Sheets(yearValue).Activate
        
        'Find number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
    'Loop through the tickers
        For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
              
    'Loop through rows in data
            Sheets(yearValue).Activate
            For j = 2 To RowCount
               '5a) Get total volume for current ticker
               If Cells(j, 1).Value = ticker Then
                   totalVolume = totalVolume + Cells(j, 8).Value
               End If
               '5b) get starting price for current ticker
               If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                   startingPrice = Cells(j, 6).Value
               End If
    
               '5c) get ending price for current ticker
               If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                   endingPrice = Cells(j, 6).Value
               End If
           Next j
            
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
        Next i
```
#### Refactored
```
'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'Create a ticker Index
    Dim tickerIndex As Integer
        tickerIndex = 0

    'Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    'Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
        
        'Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
            'Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            'Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
                
            End If
            
            'Check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
                
                ' Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                
            End If
    
        Next i
        
    Next tickerIndex
    
    ' Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
     For i = 0 To 11
     
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```

### Stock Performance Comparison
As evidenced by the results below, most of the group of stocks Styve wanted analyzed actually had decreased returns from 2017 to 2018. Only two of the twelve analyzed stocks (ENPH and RUN) produced a positive return value in 2018. One other stock (TERP) also increased it's return value in 2018, however the value is still negative. 

#### 2017
![Stock_Values_2017](https://user-images.githubusercontent.com/100883212/161404140-52c4dc21-c077-426c-bcf7-c9deb8871c10.png)

#### 2018
![Stock_Values_2018](https://user-images.githubusercontent.com/100883212/161404146-f087b7c7-5621-4b5d-8d3c-02ff57ff717b.png)

When Steve makes a stock investment recommendation to his parents, he should consider recommending only ENPH and RUN. He also might consider analyzing additional tickers to see where other positive recommendations can be made.

### Execution Time Comparison
By refactoring the code, I was able to reduce the execution time for the analysis for the 2017 stocks by ~86%.

#### 2017 Original
![VBA_Challenge_2017o](https://user-images.githubusercontent.com/100883212/161404367-36eafd9e-3aa5-4ed1-82f3-710b47f53afd.png)

#### 2017 Refactored
![VBA_Challenge_2017r](https://user-images.githubusercontent.com/100883212/161404370-2bfbdc02-d393-4193-95e0-644abb2f0c63.png)

I was also able to reduce the execution time for the analysis of the 2018 stocks by ~85%

#### 2018 Original
![VBA_Challenge_2018o](https://user-images.githubusercontent.com/100883212/161404373-f965a57c-520f-4805-a0f7-62c56a9580ea.png)

#### 2018 Refactored
![VBA_Challenge_2018r](https://user-images.githubusercontent.com/100883212/161404374-9a6e5069-b033-434e-ba62-2743f5ebbce3.png)

## Summary

