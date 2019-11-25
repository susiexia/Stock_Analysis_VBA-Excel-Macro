# Stock_Analysis_VBA-Excel-Macro
Perform stock Analysis by using Excel Macro with Visual Basic language

# Challenge

## Project Background
The purpose of this stock analysis is helping client obtain and compare *Total Daily Volume* and *Yearly Return* of each target stock in particular year.
By designing a VBA Macro for applying to different years, client can observe and make conclusion which stock performs better than others specifically in green energy industry.
Using Macro reduce the possibility of human error for repetitive calculator tasks. In addition, client can apply Macro in multiple sheets for multiple stocks simultaneously.
Refactor project to create one outer loop and three nested loops in order to loop through the stock original data only once and collect all information in a single pass.

## Conclusion
In 2018, only *ENPH* and *RUN* two stocks had positive yearly Return as well as large Total Daily Volume. Both of them was outperformance than others green stocks.

![](/2018Analysis.JPG)

In 2017, all of stocks had positive Return except *TERP* (-7.2%). "DQ" made best yearly return with 199.4% but with lowest total Daily Volume (35,796,200) in 2017.

![](/2017Analysis.JPG)


## Program Design
**There are Four Loops:**
> * (A) is the Main Loop for going through all data and assigned tickerIndex for 12 stocks respectively.
> * (B) is a nested loop in the main loop (A), go through stocks original data and retrieve ticker name, startingPrices and endingPrices, and save information to each related tickerIndex.
> * (C) a nested loop in (B) loop, in order to get volume information for each Index.
> * (D) a new loop for putting all saved output information into an analysis sheet.
## Logical Flow

1. Request users input which year they would like to analyze stock performance.
```
yearValue = InputBox("What year would you like to run the analysis on?") 
```
    
2. Create and activate an analysis worksheet to keep all information retrieved.
3. Declare 1 array for ticker and 3 outputs arrays for saving data, as well as a variable named tickerIndex.
```
    Dim tickers(12) As String
    Dim volume(12) As String
    Dim startingPrices(12) As String
    Dim endingPrices(12) As String
    
    Dim tickerIndex As Integer
```

4. Create a main loop **(A)** to assigned tickerIndex from 0 to 11. Initialize index as zero before loops.
```
    tickerIndex = 0
    For tickerIndex = 0 to 11
        if meet some criteria then
            tickerIndex = tickerIndex + 1
    Next tickerIndex
```
5. Make a **(B)** loop go through all stocks data.
```
    Worksheets(yearValue).Activate
    For J = 2 To RowCount
        If Cells(J, 1).Value <> Cells(J - 1, 1).Value Then
            tickers(tickerIndex) = Cells(J, 1).Value
            startingPrices(tickerIndex) = Cells(J, 6).Value
        End If
        
        If Cells(J + 1, 1).Value <> Cells(J, 1).Value Then
            endingPrices(tickerIndex) = Cells(J, 6).Value
            *tickerIndex = tickerIndex + 1*
            End If
    Next J    
```
6. Make a nested loop **(C)** to get incremental Daily volumes for each stock, then put into Volume(tickerIndex).
```
        For x = 2 To RowCount
            If Cells(x, 1).Value = tickers(tickerIndex) Then
                TotalVolume = TotalVolume + Cells(x, 8).Value
            End If
        Next x
     volume(tickerIndex) = TotalVolume
```
7.  Create new loop **(D)** for putting outcomes into analysis Worksheet which created on step 2
```

    Worksheets("Challenge_All Stocks Anlysis").Activate
    For i = 0 To 11      
      Cells(i + 4, 1).Value = tickers(i)
      Cells(i + 4, 3).Value = endingPrices(i) / startingPrices(i) - 1
      Cells(4 + i, 2).Value = volume(i)
    Next i
```
8. Decor Font Formatting and conditional color Formatting to analysis Worksheets
