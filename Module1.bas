Attribute VB_Name = "Module1"
Sub StockAnalysis()

    For Each ws In ThisWorkbook.Worksheets
        
        'Part 1: Column creations'
        'headers'
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Columns(10).AutoFit
        ws.Columns(11).AutoFit
        
        
        'determine the number of days in a year per stock'
        ptr = 2
        firstTicker = ws.Cells(2, 1).Value
        
        While (firstTicker = ws.Cells(ptr, 1).Value)
            ptr = ptr + 1
        Wend
        
        daysInYear = ptr - 2
        
        'loop through all stocks, skip count by the number of market days in a year'
        i = 1
        'while the ticker of the row of interest is not empty'
        While (ws.Cells((i - 1) * daysInYear + 2, 1).Value <> "")
            startRow = (i - 1) * daysInYear + 2
            endRow = i * daysInYear + 1
            
            ticker = ws.Cells(startRow, 1).Value
            
            startPrice = ws.Cells(startRow, 3).Value
            endPrice = ws.Cells(endRow, 6).Value
            
            yearlyChange = endPrice - startPrice
            percentChange = (endPrice - startPrice) / startPrice
            
            'use excels function to sum through columns set by the bounds'
            totalVolume = Application.WorksheetFunction.Sum(Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))
            
            ws.Cells(i + 1, 9).Value = ticker
            ws.Cells(i + 1, 10).Value = yearlyChange
            
            ws.Cells(i + 1, 11).Value = percentChange
            ws.Cells(i + 1, 11).NumberFormat = "0.00%"
            
            ws.Cells(i + 1, 12).Value = totalVolume
            ws.Cells(i + 1, 12).NumberFormat = "0"
            
            If yearlyChange < 0 Then
                ws.Cells(i + 1, 10).Interior.ColorIndex = 3
                ws.Cells(i + 1, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(i + 1, 10).Interior.ColorIndex = 4
                ws.Cells(i + 1, 11).Interior.ColorIndex = 4
            End If
            
            i = i + 1
        Wend
        
        'Part 2: greatest percent increase, decrease, and largest total volume'
        tickerPIncrease = ""
        greatestPIncrease = 0
        
        tickerPDecrease = ""
        greatestPDecrease = 0
        
        tickerGreatestVolume = ""
        greatestTotalVolume = 0
        
        For j = 2 To i - 1
            ticker = ws.Cells(j, 9).Value
            percentChange = ws.Cells(j, 11).Value
            totalVolume = ws.Cells(j, 12).Value
            
            If percentChange > greatestPIncrease Then
                greatestPIncrease = percentChange
                tickerPIncrease = ticker
            End If
            
            If percentChange < greatestPDecrease Then
                greatestPDecrease = percentChange
                tickerPDecrease = ticker
            End If
            
            
            If greatestTotalVolume < totalVolume Then
                greatestTotalVolume = totalVolume
                tickerGreatestVolume = ticker
            End If
        Next j
        
        'Headers'
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Columns(15).AutoFit

        
        'setting percent increase'
        ws.Cells(2, 16).Value = tickerPIncrease
        ws.Cells(2, 17).Value = greatestPIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        'setting percent decrease'
        ws.Cells(3, 16).Value = tickerPDecrease
        ws.Cells(3, 17).Value = greatestPDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        'setting volume'
        ws.Cells(4, 16).Value = tickerGreatestVolume
        ws.Cells(4, 17).Value = greatestTotalVolume
    
    Next ws
    
End Sub
