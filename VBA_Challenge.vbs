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
      Dim tickerindex As Integer
      tickerindex = 0

'1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12)  As Single
        Dim tickerEndingPrices(12) As Single

'2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(tickerindex) = 0
        
'2b) Loop over all the rows in the spreadsheet.
                 Worksheets(yearValue).Activate
                 For j = 2 To RowCount

'3a) Increase volume for current ticker


                    If Cells(j, 1).Value = tickers(tickerindex) Then
                    tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(j, 8).Value
                    End If
 
                    If Cells(j - 1, 1).Value <> tickers(tickerindex) And Cells(j, 1).Value = tickers(tickerindex) Then
                    tickerStartingPrices(tickerindex) = Cells(j, 6).Value
                    End If
 
                    If Cells(j + 1, 1).Value <> tickers(tickerindex) And Cells(j, 1).Value = tickers(tickerindex) Then
                    tickerEndingPrices(tickerindex) = Cells(j, 6).Value
                    End If
            
                Next
                
'6.Output the data for the current ticker.
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(tickerindex)
            Cells(4 + i, 2).Value = tickerVolumes(tickerindex)
            Cells(4 + i, 3).Value = tickerEndingPrices(tickerindex) / tickerStartingPrices(tickerindex) - 1
            tickerindex = tickerindex + 1
    Next
    'Next
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