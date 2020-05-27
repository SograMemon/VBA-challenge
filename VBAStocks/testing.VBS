Attribute VB_Name = "Module5"
Sub repeatNextSheet()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call forloop
    Next
    Application.ScreenUpdating = True
End Sub
Sub forloop()
    Dim i, iLast As Long
    
    Dim distTickerJ, yearlyChangeJ, percentageChangeJ, totalStockJ, percentTickerJ, percentValueJ, volJ As Integer
    Dim incI, decrI, greatestTotalI As Integer
    Dim tickerStart, tickerEnd, tickerCount, openJ, closeJ As Long
    Dim stockName As String
    Dim greatestIncValue, greatestDecrValue, greatestTotalValue As Double
    Dim greatestIncTicker, greatestDecrTicker, greatestTotalTicker As String
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    StartTime = Timer
    
    distTickerJ = 9
    yearlyChangeJ = 10
    percentageChangeJ = 11
    totalStockJ = 12
    percentTickerJ = 16
    percentValueJ = 17
    openJ = 3
    closeJ = 6
    volJ = 5
    
    incI = 2
    decrI = 3
    greatestTotalI = 4
    
    tickerCount = 2

    iLast = Range("A1").End(xlDown).Row
    'i = 2
    tickerEnd = 1
    tickerStart = 2
    
    'set lables
    Cells(1, distTickerJ).Value = "Ticker"
    Cells(1, yearlyChangeJ).Value = "yearly Change"
    Cells(1, percentageChangeJ).Value = "Percentage Change "
    Cells(1, totalStockJ).Value = "Total Stock Volume "
    
    Cells(incI, 15).Value = "Greatest % Increase"
    Cells(decrI, 15).Value = "Greatest % Decrease"
    Cells(greatestTotalI, 15).Value = "Greatest Total Volume"
    
    Cells(1, percentTickerJ).Value = "Ticker"
    Cells(1, percentValueJ).Value = "Value"
    
    Range("K:k").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    stockName = Cells(2, 1).Value
    
    For i = 2 To iLast
        tickerStart = tickerEnd + 1
        If Cells(i, 1).Value <> stockName Then
            'MsgBox (i)
            stockName = Cells(i, 1).Value
            Cells(tickerCount, distTickerJ).Value = stockName
            tickerEnd = i - 1

            Cells(tickerCount, yearlyChangeJ).Value = Cells(tickerEnd, closeJ).Value - Cells(tickerStart, openJ).Value
        
            If Cells(tickerEnd, closeJ).Value <> 0 Then
                Cells(tickerCount, percentageChangeJ).Value = (Cells(tickerEnd, closeJ).Value - Cells(tickerStart, openJ).Value) / Cells(tickerEnd, closeJ).Value
            Else
                Cells(tickerCount, percentageChangeJ).Value = (Cells(tickerEnd, closeJ).Value - Cells(tickerStart, openJ).Value)
            End If
        
            Cells(tickerCount, totalStockJ).Value = WorksheetFunction.Sum(Range(Cells(tickerStart, volJ), Cells(tickerEnd, volJ)))
        
            If Cells(tickerCount, percentageChangeJ).Value < 0 Then
                Cells(tickerCount, percentageChangeJ).Interior.ColorIndex = 3
            Else
                Cells(tickerCount, percentageChangeJ).Interior.ColorIndex = 4
            End If
        
            If i = 2 Then
                greatestIncTicker = stockName
                greatestDecrTicker = stockName
                greatestTotalTicker = stockName
            
                greatestIncValue = Cells(tickerCount, percentageChangeJ).Value
                greatestDecrValue = Cells(tickerCount, percentageChangeJ).Value
                greatestTotalValue = Cells(tickerCount, totalStockJ).Value
        Else
                If greatestIncValue < Cells(tickerCount, percentageChangeJ).Value Then
                    greatestIncTicker = stockName
                    greatestIncValue = Cells(tickerCount, percentageChangeJ).Value
                End If
                If greatestDecrValue > Cells(tickerCount, percentageChangeJ).Value Then
                    greatestDecrTicker = stockName
                    greatestDecrValue = Cells(tickerCount, percentageChangeJ).Value
                End If
                If greatestTotalValue < Cells(tickerCount, totalStockJ).Value Then
                    greatestTotalTicker = stockName
                    greatestTotalValue = Cells(tickerCount, totalStockJ).Value
                End If
            End If
            tickerCount = tickerCount + 1
        End If
        
    Next i
    
    Cells(incI, percentTickerJ).Value = greatestIncTicker
    Cells(decrI, percentTickerJ).Value = greatestDecrTicker
    Cells(greatestTotalI, percentTickerJ).Value = greatestTotalTicker
    
    Cells(incI, percentValueJ).Value = greatestIncValue
    Cells(decrI, percentValueJ).Value = greatestDecrValue
    Cells(greatestTotalI, percentValueJ).Value = greatestTotalValue
    
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)

    'Notify user in seconds
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

    
End Sub


