Sub AllStocksAnalysis()
    
    'initialize our time variables track how long the code takes to execute
    Dim startTime As Single
    Dim endTime As Single

    'input by the user for the year (either 2017 or 2018)
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'create header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'Initialize array of all tickers
    Dim tickers(11) As String
    
    'Assigning each element of our array to a ticker
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
    RowCount = Cells.Find(What:="*", SearchDirection:=xlPrevious).Row
    
    'create a tickerIndex
    tickerIndex = 0
    
    'create three output arrays
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    Dim tickerVolumes(11) As Long
    
    
    'initialize the tickerVolumes to zero
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
        
    Worksheets(yearValue).Activate
    For i = 2 To RowCount
        
        ticker = tickers(tickerIndex)
            
        'Get total volume for current ticker
        If Cells(i, 1).Value = ticker Then
                
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        End If
            
        'Get starting price for current ticker
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
            
        'Get ending price for current ticker
        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
                
        End If
            
    Next i
        
          
    For i = 0 To 11
    
        'Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
        
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    
    
    
    dataRowStart = 4
    dataRowEnd = 15
    
    'Giving our analysis output colors to indicate whether the stocks did well
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
    
    'output to show the time it took for our code to finish running
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub