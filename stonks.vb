


Sub calculate()

For s = 1 To Sheets.Count
Dim ws As Worksheet
Set ws = Worksheets(s)

ws.Cells(1, 10) = "Ticker"
ws.Cells(1, 11) = "Amount Change"
ws.Cells(1, 12) = "Percentage Change"
ws.Cells(1, 13) = "Stock Volume"

ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Stock Volume"

greatestincrease = 0
greatestdecrease = 0
maxvolume = 0
usedrows = ws.Rows.Count
outputrow = 2

For i = 2 To usedrows - 1
    ticker = ws.Cells(i, 1)
    nexttick = ws.Cells(i + 1, 1)
    If i = usedrows - 1 Then
            stvolume = ws.Cells(i, 7) + stvolume
            openvalue = ws.Cells(i + 1, 3)
            high = ws.Cells(i + 1, 4)
            low = ws.Cells(i + 1, 5)
            closevalue = ws.Cells(i + 1, 6)
            stvolume = ws.Cells(i + 1, 7) + stvolume
            ws.Cells(outputrow, 10) = ticker
            ws.Cells(outputrow, 11) = closevalue - initialopen
            If initialopen = 0 Then
            
            Else
            percentagechange = ((closevalue - initialopen) / initialopen) * 100
            ws.Cells(outputrow, 12) = STR(percentagechange)+" %"
            End If
            ws.Cells(outputrow, 13) = stvolume
    ElseIf i = 2 Then
        initialopen = ws.Cells(i, 3)
        stvolume = ws.Cells(i, 7)
    Else
        If ticker = nexttick Then
            stvolume = stvolume + ws.Cells(i, 7)
        Else
            openvalue = ws.Cells(i, 3)
            high = ws.Cells(i, 4)
            low = ws.Cells(i, 5)
            closevalue = ws.Cells(i, 6)
            stvolume = ws.Cells(i, 7) + stvolume
            ws.Cells(outputrow, 10) = ticker
            ws.Cells(outputrow, 11) = closevalue - initialopen
            If initialopen = 0 Then
        
            Else
            percentagechange = ((closevalue - initialopen) / initialopen) * 100
            ws.Cells(outputrow, 12) = STR(percentagechange)+" %"
                If ws.Cells(outputrow, 12) > 0 Then
                    ws.Cells(outputrow, 12).Interior.ColorIndex = 4
                ElseIf ws.Cells(outputrow, 12) < 0 Then
                    ws.Cells(outputrow, 12).Interior.ColorIndex = 3
                Else
                    ws.Cells(outputrow, 12).Interior.ColorIndex = 6
                End If
            End If
            ws.Cells(outputrow, 13) = stvolume
            initialopen = ws.Cells(i + 1, 3)
            If stvolume > maxvolume Then
                maxvolume = stvolume
                ws.Cells(4, 16) = ticker
                ws.Cells(4, 17) = maxvolume
            End If
            If percentagechange > greatestincrease Then
                greatestincrease = percentagechange
                ws.Cells(2, 16) = ticker
                ws.Cells(2, 17) = Str(greatestincrease) + " %"
            End If
            If percentagechange < greatestdecrease Then
                greatestdecrease = percentagechange
                ws.Cells(3, 16) = ticker
                ws.Cells(3, 17) = Str(greatestdecrease) + " %"
            End If
            outputrow = outputrow + 1
            stvolume = ws.Cells(i + 1, 7)
        End If
    End If
Next i
Next s

End Sub



