Attribute VB_Name = "Module2"
Sub repeatNextSheet()
'This sub applies the same macro to all sheets in this excel file
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call forloop 'This is the macro that will be applied to each sheet
    Next
    Application.ScreenUpdating = True
End Sub
Sub forloop()
'This prints the column names and performs the calculations for yearly change, percentage change, total stock volume for each of the stocks.
'It also calculates and prints greatest % increase, decrease and the greatest total volume of the stocks

    'Variable Declaration
    
    'row pointers for start and end of for loop
    Dim i, iLast As Long
    
    'Column pointers for each colum name to be generated
    Dim distTickerJ, yearlyChangeJ, percentageChangeJ, totalStockJ, percentTickerJ, percentValueJ, volJ As Integer
    
    Dim incI, decrI, greatestTotalI As Integer
    Dim tickerStart, tickerEnd, tickerCount, openJ, closeJ As Long
    Dim stockName As String
    Dim greatestIncValue, greatestDecrValue, greatestTotalValue As Double
    Dim greatestIncTicker, greatestDecrTicker, greatestTotalTicker As String
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    'start time to calculate execution time
    StartTime = Timer
    
    'column index assignment for each column name
    distTickerJ = 9
    yearlyChangeJ = 10
    percentageChangeJ = 11
    totalStockJ = 12
    percentTickerJ = 16
    percentValueJ = 17
    openJ = 3
    closeJ = 6
    volJ = 5
    
    'row index assignment for each row name
    incI = 2
    decrI = 3
    greatestTotalI = 4
    
    'row index to keep track of how many distict tickers calculated
    tickerCount = 2
    
    'the index of the last row
    iLast = Range("A1").End(xlDown).Row
    
    'tickerStart is the start of a distict ticker and tickerEnd is the end index of that distinct ticker
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
    
    'formate cells as percentage
    Range("K:k").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    'intialize stockName as the first ticker value
    stockName = Cells(2, 1).Value
    
    For i = 2 To iLast
        'move to next ticker by adding 1 to the the tickerEnd of last ticker
        tickerStart = tickerEnd + 1
        
        'Continue incrementing i till current ticker not equil to ticker in stockName Variable
        If Cells(i, 1).Value <> stockName Then
            'Set stockName to new ticker value for next comparison
            stockName = Cells(i, 1).Value
            'Print distinct ticker
            Cells(tickerCount, distTickerJ).Value = Cells(tickerStart, 1).Value
            'set tickerEnd as current location - 1
            tickerEnd = i - 1

            'Calculate yearly change and print it
            Cells(tickerCount, yearlyChangeJ).Value = Cells(tickerEnd, closeJ).Value - Cells(tickerStart, openJ).Value
        
            If Cells(tickerStart, openJ).Value <> 0 Then 'Avoid division by zero
                'Calculate % yearly change when oppening price is not zero
                Cells(tickerCount, percentageChangeJ).Value = (Cells(tickerEnd, closeJ).Value - Cells(tickerStart, openJ).Value) / Cells(tickerStart, openJ).Value
            Else
                Cells(tickerCount, percentageChangeJ).Value = (Cells(tickerEnd, closeJ).Value - Cells(tickerStart, openJ).Value)
            End If
            'Calculate sum of all rows in <vol> from index values tickerStart to tickerEnd. Print this value in total stock value
            Cells(tickerCount, totalStockJ).Value = WorksheetFunction.Sum(Range(Cells(tickerStart, volJ), Cells(tickerEnd, volJ)))
            
            'Cell color formating
            If Cells(tickerCount, percentageChangeJ).Value < 0 Then
                Cells(tickerCount, percentageChangeJ).Interior.ColorIndex = 3 ' red color if negative value
            Else
                Cells(tickerCount, percentageChangeJ).Interior.ColorIndex = 4 ' green color if positive value
            End If
        
            If i = 2 Then 'initialize first stock value as the greatest increase and decrease
                greatestIncTicker = stockName
                greatestDecrTicker = stockName
                greatestTotalTicker = stockName
            
                greatestIncValue = Cells(tickerCount, percentageChangeJ).Value
                greatestDecrValue = Cells(tickerCount, percentageChangeJ).Value
                greatestTotalValue = Cells(tickerCount, totalStockJ).Value
            Else
                If greatestIncValue < Cells(tickerCount, percentageChangeJ).Value Then ' for each new stock check if it is more than curent greatest increase
                    greatestIncTicker = stockName ' if new stock greater than current greatest assign new stock to the gratest increase variables
                    greatestIncValue = Cells(tickerCount, percentageChangeJ).Value
                End If
                If greatestDecrValue > Cells(tickerCount, percentageChangeJ).Value Then ' for each new stock check if it is less than curent greatest decrease
                    greatestDecrTicker = stockName ' if new stock less than current greatest assign new stock to the gratest decrease variables
                    greatestDecrValue = Cells(tickerCount, percentageChangeJ).Value
                End If
                If greatestTotalValue < Cells(tickerCount, totalStockJ).Value Then  ' for each new stock check if it is more than curent greatest total
                    greatestTotalTicker = stockName ' if new stock greater than current greatest assign new stock to the gratest total variables
                    greatestTotalValue = Cells(tickerCount, totalStockJ).Value
                End If
            End If
            'increment to print next distinct ticker to next row
            tickerCount = tickerCount + 1
        End If
        
    Next i
    
    'print greatest increase, decrease and total ticker names
    Cells(incI, percentTickerJ).Value = greatestIncTicker
    Cells(decrI, percentTickerJ).Value = greatestDecrTicker
    Cells(greatestTotalI, percentTickerJ).Value = greatestTotalTicker
    
    'print greatest increase, decrease and total ticker values
    Cells(incI, percentValueJ).Value = greatestIncValue
    Cells(decrI, percentValueJ).Value = greatestDecrValue
    Cells(greatestTotalI, percentValueJ).Value = greatestTotalValue
    
    'Determine how many seconds code took to run
    SecondsElapsed = Round(Timer - StartTime, 2)

    'Notify user in seconds the execution time
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

    
End Sub
