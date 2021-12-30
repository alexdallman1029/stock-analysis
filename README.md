# Green Stocks Analysis
## Overview
### Background
A client, Steve, is interested in obtaining stock information to make an informed decision about what "green" stocks to invest in for his parents based on the total daily volume and return on the stock for a given year. He wants to expand an existing dataset of 12 stocks to include the entire stock market over the last few years. 
### Purpose
The purpose of this analysis is to refactor exsiting code to make analyzing stock outcome by year more efficient and applicable to an expanded data set using VBA code in Excel. 
## Results
### Method
To refactor, the variable <b>tickerIndex</b> was set equal to zero, iterated over all the rows, and used to access the correct index across four different arrays, including the tickers array, tickerVolumes array, tickerStartingPrices array, and tickerEndingPrices array.


The VBA Script is as follows: 
    
    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    Dim tickerIndex As Single
    tickerIndex = 0

    '1b) Create three output arrays
    
    ReDim tickerVolumes(12) As Long
    ReDim tickerStartingPrices(12) As Single
    ReDim tickerEndingPrices(12) As Single
    
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
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        End If
        'End If
        
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
     
     tickerIndex = i
     
     Cells(i + 4, 1).Value = tickers(tickerIndex)
     Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
     Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
    
    Next i
    
    'Formatting
    
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub

### Analysis
The refactored code worked more efficiently than the original code and is now more adaptable for a new, larger data set of stocks to analyze. Figures 1 and 2 show the original runtime for the code without setting a variable across arrays. Figures 3 and 4 show the refactored code. 

<br><b>Figure 1</b><br><br>
<img src="stock-analysis/resources/VBA_Challenge_Orignial_2017.png" width=500><br>

<br><b>Figure 2</b><br><br>
<img src="stock-analysis/resources/VBA_Challenge_Orignial_2018.png" width=500><br>

<br><b>Figure 3</b><br><br>
<img src="stock-analysis//VBA_Challenge_2017.png" width=500><br>

<br><b>Figure 4</b><br><br>
<img src="stock-analysis//VBA_Challenge_2018.png" width=500><br>

## Summary
### Advantages of Refactoring Code
### Disadvantages of Refactoring Code
### Application of Pros and Cons to Refactoring Original VBA Script

