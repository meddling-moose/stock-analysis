Attribute VB_Name = "Module1"
Sub tickerLoop()
'Initialize variables
Dim ticker, nextTicker, prevTicker, gVTicker, gPITicker, gPDTicker As String
Dim row As Long
Dim opPrice, edPrice, gPIncrease, gPDecrease, percentChange, yearlyChange As Double
Dim volume, gVolume As LongLong
Dim tickerCount As Integer
Dim current As Worksheet

tickerCount = 1
gVolume = 0
gPIncrease = 0
gPDecrease = 0

'Loop through all of the sheets at once
For Each ws In ActiveWorkbook.Worksheets

    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"
    'Loop through the ticker rows
    For row = 2 To Cells(Rows.Count, 1).End(xlUp).row 'Find the last row
        'Find when the ticker is about to change
        ticker = ws.Cells(row, 1).Value
        nextTicker = ws.Cells(row + 1, 1).Value
        prevTicker = ws.Cells(row - 1, 1).Value
        
        'Condition on first of a ticker run, middle, or end (about to change)
        If prevTicker <> ticker Then 'we have found the first of a ticker
            'if beginning of a run, output the ticker, and record opening price and volume
            volume = ws.Cells(row, 7).Value
            opPrice = ws.Cells(row, 3).Value
            tickerCount = tickerCount + 1
        
        ElseIf (ticker = nextTicker) And (prevTicker = ticker) Then 'we have found a part of the run
            'if the middle of a run, continue summing the volume
            volume = volume + ws.Cells(row, 7).Value
        
        ElseIf ticker <> nextTicker Then 'we have found the end of a run
            'if end of a run, record closing price, then calculate yearly change, percent change, and conclude volume and output data
            edPrice = ws.Cells(row, 6).Value
            volume = volume + ws.Cells(row, 7).Value
            
            ws.Cells(tickerCount, 9) = ticker
            
            yearlyChange = edPrice - opPrice
            ws.Cells(tickerCount, 10) = yearlyChange
            'conditional format here to have the background be green or red depending on greater than or less than zero
            If yearlyChange >= 0 Then
                'make the background green
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 4
            ElseIf yearlyChange < 0 Then
                'make the background of the cell red
                ws.Cells(tickerCount, 10).Interior.ColorIndex = 3
            End If
            
            percentChange = (edPrice / opPrice) - 1
            ws.Cells(tickerCount, 11) = percentChange
            'conditional formatting here to show appropriate percentages
            ws.Cells(tickerCount, 11).Value = Format(percentChange, "#.##%")
            
            'compare to greatest percent increase and greatest percent decrease, and if it exceeds one of them, replace the logged info with this new one
            If ((percentChange > 0) And (percentChange > gPIncrease)) Or ((percentChange < 0) And (percentChange < gPDecrease)) Then
                If percentChange > 0 Then
                    gPIncrease = percentChange
                    gPITicker = ticker
                Else
                    gPDecrease = percentChange
                    gPDTicker = ticker
                End If
            End If
            
            Cells(tickerCount, 12) = volume
            'Compare this volume to greatest recorded volume and if bigger, record it
            If volume > gVolume Then
                gVTicker = ticker
                gVolume = volume
            End If
        End If
    Next row
    'Output the greatest percent increase, decrease and total volume information
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    'Output variable information
    ws.Cells(2, 16) = gPITicker
    ws.Cells(2, 17) = gPIncrease
    ws.Cells(2, 17).Value = Format(gPIncrease, "#.##%")
    ws.Cells(3, 16) = gPDTicker
    ws.Cells(3, 17) = gPDecrease
    ws.Cells(3, 17).Value = Format(gPDecrease, "#.##%")
    ws.Cells(4, 16) = gVTicker
    ws.Cells(4, 17) = gVolume
    
    'reset counter variables
    tickerCount = 1
    gVolume = 0
    gPIncrease = 0
    gPDecrease = 0
Next

End Sub

