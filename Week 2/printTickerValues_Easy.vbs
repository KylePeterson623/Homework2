Sub printTickerValues():

Dim ticker As String

Dim tickerVolume As Double
    tickerVolume = 0

Dim tickerLocation As Integer
    tickerLocation = 2

For Each ws In Worksheets
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

For i = 2 To lastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    tickerName = Cells(i, 1).Value
    tickerVolume = tickerVolume + Cells(i, 7).Value
    Range("I" & tickerLocation).Value = tickerName
    Range("J" & tickerLocation).Value = tickerVolume
    tickerLocation = tickerLocation + 1
    tickerVolume = 0
Else
    tickerVolume = tickerVolume + Cells(i, 7).Value
End If

Next i
Next ws

End Sub