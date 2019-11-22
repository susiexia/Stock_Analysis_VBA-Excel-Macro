Sub HW_AllStocksAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("Challenge_All Stocks Anlysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'declare 4 arrays
    Dim tickers(12) As String
    Dim volume(12) As String
    Dim startingPrices(12) As String
    Dim endingPrices(12) As String
    'create index variable
    Dim tickerIndex As Integer

    Worksheets(yearValue).Activate
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '(1)the outer loop for index from 0 to 11
    tickerIndex = 0
    
    Worksheets(yearValue).Activate
    For tickerIndex = 0 To 11
        '(2)the main loop for stock data
        Worksheets(yearValue).Activate
        For J = 2 To RowCount
            'retrieve ticker name and start price for each tickerIndex and store them in arrays
            If Cells(J, 1).Value <> Cells(J - 1, 1).Value Then
                tickers(tickerIndex) = Cells(J, 1).Value
                startingPrices(tickerIndex) = Cells(J, 6).Value
            End If
                '(3)a nested loop for retrieving TotalVolume for each volume array
                Worksheets(yearValue).Activate
                    TotalVolume = 0
                    For x = 2 To RowCount
                        If Cells(x, 1).Value = tickers(tickerIndex) Then
                            TotalVolume = TotalVolume + Cells(x, 8).Value
                        End If
                    Next x

                    volume(tickerIndex) = TotalVolume
            
            'retrieve and store ending price in array as well as increment tickerIndex for next loop
            If Cells(J + 1, 1).Value <> Cells(J, 1).Value Then
                endingPrices(tickerIndex) = Cells(J, 6).Value
                tickerIndex = tickerIndex + 1
            End If
        Next J
        
    Next tickerIndex
    
    '(4)store all informations collected in a output worksheet
    Worksheets("Challenge_All Stocks Anlysis").Activate
    For i = 0 To 11
        
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 3).Value = endingPrices(i) / startingPrices(i) - 1
        Cells(4 + i, 2).Value = volume(i)
    
    Next i
                
    'formatting
    Worksheets("Challenge_All Stocks Anlysis").Activate
        Range("A3:C3").Font.Bold = True
        Range("A1").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("c4:c15").NumberFormat = "0.0%"
        Columns(2).AutoFit
    'color conditional formatting
    Worksheets("Challenge_All Stocks Anlysis").Activate
    dataRowEnd = Cells(Rows.Count, "C").End(xlUp).Row
    dataRowStart = 4
    For r = dataRowStart To dataRowEnd
        If Cells(r, 3).Value > 0 Then
            Cells(r, 3).Interior.Color = vbGreen
        ElseIf Cells(r, 3).Value < 0 Then
            Cells(r, 3).Interior.Color = vbRed
        Else
            Cells(r, 3).Interior.Color = xlNone
        End If
    Next r


End Sub
