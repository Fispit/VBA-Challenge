


Sub calculate()

For s = 1 To Sheets.Count
Dim ws As Worksheet

Set ws = Worksheets(s)

ws.Cells(1, 10) = "Ticker"
ws.Cells(1, 11) = "Amount Change"
ws.Cells(1, 12) = "Percentage Change"
ws.Cells(1, 13) = "Stock Volume"

ws.Cells(1, 14) = "Ticker"
ws.Cells(1, 15) = "Amount Change"
ws.Cells(1, 16) = "Percentage Change"
ws.Cells(1, 17) = "Stock Volume"
ws.Cells(1, 18) = "Stock Volume"


usedrows = Rows.Count
MsgBox (Str(usedrows))
outputrow = 2
outputrow2 = 2
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
            ws.Cells(outputrow, 12) = (closevalue - initialopen) / initialopen
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
            ws.Cells(outputrow, 12) = ((closevalue - initialopen) / initialopen) * 100
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
            outputrow = outputrow + 1
            stvolume = ws.Cells(i + 1, 7)




        End If




    End If




Next i
Next s

End Sub


