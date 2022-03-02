# Stock Analysis with VBA

## Overview of Project
The purpose of the project is to easily summarize the 12 different stocks available (Under A - Column) with each individuals stocks’ yearly volume (Under B – Column), and the yearly closing price percentage (Under C – Column). These three key markers can give Steve’s family a better idea of what stocks to invest in.
The purpose of this classroom project however, is to increase efficiency and reduce the time to run the code.


#### Data
Total Daily Volume
Under Steve’s family view, the more a stock is traded, the better performance it will have. By summing up each stocks daily trading amount and filtering it by year, Steve can see the amount of yearly activity the stock is moving. 
#### Return 
A good indicator of a healthy stock is seeing an increase in stock price at the end of the year. By subtracting each individual stocks’ earliest closing price by year and the last closing price, Steve will be able to view which stocks have increased in price and which have decreased (in percentage).

## Results
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
The script from the original file and the challenge were almost identical except that the new file has a few changes, which were easily transferrable. 
Running the code however, things seemed to go much smoother with the challenge VBA script is about  .1 sec faster than the original script.



![Orig_VBA Code_ 2017   2018](https://user-images.githubusercontent.com/98041751/156438116-40b05dfd-1706-461c-b867-58cd653190d3.png)

_**Refectored Code**_

![VBA_Challenge_2017](https://user-images.githubusercontent.com/98041751/156438295-d0db4f4d-cd23-4e22-8521-7bc15366efd8.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/98041751/156438307-d3e3a62b-8b37-4510-9b02-c2d8de4c10db.png)



## Summary: 
The main advantage to using refactoring code is the time efficiency and organization of the code. Which is helpful for finding errors and debugging. It can also allow others to read your code much more easily. However, if the code is large, complex and with many moving parts, refactoring may not be the most practical solution, especially if there is already pre-existing code.  



## How do these pros and cons apply to refactoring the original VBA script?
I did experience a better time efficiency after refactoring, even if it was only .1 sec better. I believe from the original code, it was easier to read as well, as seen below. Another pro on refactoring is that you can use as much of the relevant old code with simply copying and pasting. 



    
    
    Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
      
    range("A1").Value = "All Stocks (" + yearValue + ")"
        
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
    RowCount = Cells(rows.Count, "A").End(xlUp).Row
        
    '1a) Create a ticker Index
    Ticker = tickers(i)
    
    '1b) Create three output arrays
        
    Dim tickerVolumes As Long
    Dim tickerSartingPrices As Single
    Dim tickerEndingPrices As Single
     
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
      For i = 0 To 11
        totalVolumes = 0
        Ticker = tickers(i)
        Worksheets(yearValue).Activate
        
        
            ''2b) Loop over all the rows in the spreadsheet.
            For j = 2 To RowCount
            
                '3a) Increase volume for current ticker
                If Cells(j, 1).Value = Ticker Then
                    totalVolumes = totalVolumes + Cells(j, 8).Value
                End If
                
                '3b) Check if the current row is the first row with the selected tickerIndex.
                'If  Then
                If Cells(j - 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
                    tickerSartingPrices = Cells(j, 6).Value
                End If
                'End If
                
                '3c) check if the current row is the last row with the selected ticker
                'If the next row's ticker doesn't match, increase the tickerIndex.
                'If  Then
                
                If Cells(j + 1, 1).Value <> Ticker And Cells(j, 1).Value = Ticker Then
                    tickerEndingPrices = Cells(j, 6).Value
                End If
                'End If
            
            Next j
        
        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = Ticker
        Cells(4 + i, 2).Value = totalVolumes
        Cells(4 + i, 3).Value = tickerEndingPrices / tickerSartingPrices - 1
        
    Next i
      
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    range("A3:C3").Font.FontStyle = "Bold"
    range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    range("B4:B15").NumberFormat = "#,##0"
    range("C4:C15").NumberFormat = "0.0%"
    
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


