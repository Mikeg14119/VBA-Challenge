Sub MarketTicker()
    
    ' Define variables for loop
    Dim lastRow As Long
    Dim Ticker As String
    Dim tickerStart As Long
    Dim tickerEnd As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    

    'Loop through all worksheets with stock data
    For Each ws In Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            'Determine last row
        
        'Create output headers for summary info
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        
        'j will be used for the summary table row
        j = 1
        
        'Loop through all tickers in current worksheet
        For i = 2 To lastRow
            'Used to find the first row and initialize ticker start variables
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                tickerStart = i
    
                    
                'Calculate summary info for ticker
                Ticker = ws.Cells(tickerStart, 1).Value
                openPrice = ws.Cells(tickerStart, 3).Value
            End If
            
            'Used to find the final row and calculate all the full year stats
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                tickerEnd = i

                ' Calculate summary info for ticker
                Ticker = ws.Cells(tickerStart, 1).Value
                openPrice = ws.Cells(tickerStart, 3).Value
                If IsNumeric(ws.Cells(tickerEnd + 1, 6).Value) Then
                    closePrice = CDbl(ws.Cells(tickerEnd, 6).Value)
                    yearlyChange = closePrice - openPrice
                    If openPrice <> 0 Then
                        percentChange = yearlyChange / openPrice
                    Else
                        percentChange = 0
                    End If
                totalVolume = WorksheetFunction.Sum(Range("G" & tickerStart & ":G" & tickerEnd))

        
                End If
                
                ' Output summary information for ticker
                
                ws.Range("I" & (j + 1)).Value = Ticker
                ws.Range("J" & (j + 1)).Value = yearlyChange
                ws.Range("K" & (j + 1)).Value = percentChange
                ws.Range("L" & (j + 1)).Value = totalVolume
                
                'Apply conditional formatting to yearly change
                Set formatRange = ws.Range("K" & (j + 1))
                If yearlyChange > 0 Then
                    formatRange.Interior.Color = RGB(0, 255, 0) ' Green
                ElseIf yearlyChange < 0 Then
                    formatRange.Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                'Increment J
                j = j + 1
                
            End If
                                                                             
        Next i
    Next ws
      
End Sub