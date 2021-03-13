
Sub calculate()
Cells(1, 10) = "Ticker"
Cells(1, 11) = "Amount Change"
Cells(1, 12) = "Percentage Change"
Cells(1, 13) = "Stock Volume"

usedrows = Rows.Count 'Found online on how to find how many rows are in a column
MsgBox (Str(usedrows))
outputrow = 2
For i = 2 To usedrows - 1
    ticker = Cells(i, 1)
    nexttick = Cells(i + 1, 1)
    If i = usedrows - 1 Then
            stvolume = Cells(i, 7) + stvolume
            openvalue = Cells(i + 1, 3)
            high = Cells(i + 1, 4)
            low = Cells(i + 1, 5)
            closevalue = Cells(i + 1, 6)
            stvolume = Cells(i + 1, 7) + stvolume
            Cells(outputrow, 10) = ticker
            Cells(outputrow, 11) = closevalue - initialopen
            If initialopen = 0 Then
            
            Else
            Cells(outputrow, 12) = (closevalue - initialopen) / initialopen
            End If
            Cells(outputrow, 13) = stvolume
            
    ElseIf i = 2 Then
        initialopen = Cells(i, 3)
        stvolume = Cells(i, 7)
    Else

        If ticker = nexttick Then
            stvolume = stvolume + Cells(i, 7)

        
        Else
            openvalue = Cells(i, 3)
            high = Cells(i, 4)
            low = Cells(i, 5)
            closevalue = Cells(i, 6)
            stvolume = Cells(i, 7) + stvolume
            Cells(outputrow, 10) = ticker
            Cells(outputrow, 11) = closevalue - initialopen
            If initialopen = 0 Then
            
            Else
            Cells(outputrow, 12) = (closevalue - initialopen) / initialopen
            End If
            
            Cells(outputrow, 13) = stvolume
            initialopen = Cells(i + 1, 3)
            outputrow = outputrow + 1
            stvolume = Cells(i + 1, 7)




        End If




    End If




Next i

End Sub

