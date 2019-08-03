Sub tickerCode()

Cells(1, "I").Value = "Ticker"
Cells(1, "J").Value = "Total Stock Volume"

Dim Ticker as String
Dim Volume as Double
Dim Row as Double
Row = 2
Dim Column as Integer
Column = 1
Dim i as Long

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 to lastrow

    If Cells(i+1, Column).Value <> Cells(i,Column).Value Then
        Ticker = Cells(i, Column).Value
        Cells(Row, Column + 8).Value = Ticker
        Volume = Volume + Cells(i, Column + 6).Value
            Cells(Row, Column + 9).Value = Volume
                Row = Row + 1
                Volume = 0
    Else
        Volume = Volume + Cells(i, Column + 6).Value
    End If
            
Next i

End Sub