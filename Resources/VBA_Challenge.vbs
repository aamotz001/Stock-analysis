'A speed optimized subroutine which calculates stock parameters
Sub AllStocksAnalysisRefactored()
    
    'Declare variables used to calculate the time the subroutine takes
    Dim startTime As Single
    Dim endTime  As Single

    'Get the desired year value based on user input
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Get the time on the clock immediately after the user input
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'Set sheet title
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
            
    '1a) Create a ticker Index and set equal to zero
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        'Set ticker volume array = 0 by initializing each
        tickerVolumes(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            'Set starting price for current ticker
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next row's ticker doesn't match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            'Set ending price for current ticker
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        'Select the correct worksheet
        Worksheets("All Stocks Analysis").Activate
        
        'Set the calculated values in the worksheet
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    
    'Make headers bold
    Range("A3:C3").Font.FontStyle = "Bold"
    'Add a line under headers
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Format the TDV values as a dollar
    Range("B4:B15").NumberFormat = "#,##0"
    'Format the return as a perent
    Range("C4:C15").NumberFormat = "0.0%"
    'Set the column widths equal to the widest value
    Columns("B").AutoFit
    
    
    'Set variables equal to the start and stop row index for the cells containing data
    dataRowStart = 4
    dataRowEnd = 15
    
    'Loop over the cells containing the data
    For i = dataRowStart To dataRowEnd
    
        'Set cell color equal to green if return was positive
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        'Set cell color equal to red if return was negative
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    
    'Get the time on the clock once the loops, data input and formatting have finished
    endTime = Timer
    
    'Calculate and output code run time to user
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub