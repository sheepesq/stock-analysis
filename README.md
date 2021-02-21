# stock-analysis

## Purpose

The purpose of this project was to make an efficient way to look at the stocks provided using the vba code we have learned over the past week. Our goal was to simplify (refactor) the code to make our analysis of the date more efficient. 


## Results

### Attached pictures
[2017 time results] (https://github.com/sheepesq/stock-analysis/blob/main/VBA_Challenge_2017.png)


[2018 time results] (https://github.com/sheepesq/stock-analysis/blob/main/VBA_Challenge_2018.png)

### Refactorted code

'1a) Create a ticker Index
    Dim tickerIndex As Single
    
    tickerIndex = 0
    
    '1b) Create three output arrays
        Dim tickervolume(12) As Long
        Dim tickerstartprice(12) As Single
        Dim tickerendprice(12) As Single
        
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
       For i = 0 To 11
        tickervolume(i) = 0
      Next i
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        '3a) Increase volume for current ticker
            tickervolume(tickerIndex) = tickervolume(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerstartprice(tickerIndex) = Cells(i, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
       
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerendprice(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
            End If
    
        Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        ' tickerIndex = i
        'display data
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickervolume(i)
        Cells(i + 4, 3).Value = tickerendprice(i) / tickerstartprice(i) - 1
       

## Summary 

### Pros and Cons of Refactoring

The upside of refactoring code is that you are streamlining what needs to be done on and increasing the efficency of the program which could greatly reduce the run time on large projects. The downside to refactoring code is that you are not leaving good enough alone, in otherwords, you are taking something that works and spending time and effort to make it work better.

### VBA Refactoring pros and cons

An upside of refactoring in VBA  is that you have portions of the code already created for you which will cut down on the time spent refactoring. A downside is that very much like the con of refactorting in general  is that you can be looking to simplify something which is beyond your skillset and will have wasted time working on something that already works. 
