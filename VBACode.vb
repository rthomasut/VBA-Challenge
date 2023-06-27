Sub StockAnalysis()
    'Set variables
    Dim ws As Worksheet
    Dim ticker As Double
    Dim firstrow As Double
    Dim currentrow As Double
    Dim lastrow As Double
    Dim openyear As Double
    Dim closingyear As Double
    Dim totalvolume As Double

    
    For Each ws In Sheets
        Worksheets(ws.Name).Activate
        ticker = 2
        firstrow = 2
        currentrow = 2
        lastrow = WorksheetFunction.CountA(ActiveSheet.Columns(1))
        totalvolume = 0
        
        'Label output rows and cell values
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "% Change"
        Cells(1, 12).Value = "Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
    'Find and print unique tickers in "Ticker" column
        For i = firstrow To lastrow
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i - 1, 1).Value
            If tickers <> tickers2 Then
                Cells(currentrow, 9).Value = tickers
                currentrow = currentrow + 1
            End If
         Next i
         
        'Summarize each tickers volume
        For i = firstrow To lastrow + 1
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i - 1, 1).Value
            'Summarize each unique ticker volume
            If tickers = tickers2 And i > 2 Then
                totalvolume = totalvolume + Cells(i, 7).Value
            'if ticker changes move down one row and continue
            ElseIf i > 2 Then
                Cells(ticker, 12).Value = totalvolume
                ticker = ticker + 1
                totalvolume = 0
            Else
                totalvolume = totalvolume + Cells(i, 7).Value
            End If
        Next i
        
        'Get first and last year of all tickers
        ticker = 2
        For i = firstrow To lastrow
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                closingyear = Cells(i, 6).Value
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                openyear = Cells(i, 3).Value
            End If
            'Ensure there is a difference between openyear and closeyear
            If openyear > 0 And closingyear > 0 Then
                yearchange = closingyear - openyear
                peryearchange = yearchange / openyear
                'limit the output to 2 decimal places
                Cells(ticker, 10).Value = FormatNumber(yearchange)
                Cells(ticker, 11).Value = FormatPercent(peryearchange)
                closingyear = 0
                openyear = 0
                ticker = ticker + 1
            End If
        Next i
        
        'Finds min and max values, then assigns each value to proper cell
        MaxPercent = WorksheetFunction.Max(ActiveSheet.Columns("k"))
        MinPercent = WorksheetFunction.Min(ActiveSheet.Columns("k"))
        maxvolume = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        
        Range("Q2").Value = FormatPercent(MaxPercent)
        Range("Q3").Value = FormatPercent(MinPercent)
        Range("Q4").Value = maxvolume
        
        
        'Loops through columns "K - % Change" and "L - Total Volume", find and print ticker and value
        For i = firstrow To lastrow
            If MaxPercent = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
            ElseIf MinPercent = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
            ElseIf maxvolume = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            End If
        Next i
        
        'Check each value in column "J - Yearly Change"
        'Highlight positive values green and negative values red
        For i = firstrow To lastrow
            If IsEmpty(Cells(i, 10).Value) Then Exit For
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    Next ws
                
End Sub

