Sub Easy()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = " Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        
        Sum = 0
        Result_row = 2
        For i = 2 To LastRow
            NextTicker = Cells(i + 1, 1).Value
            If (Cells(i, 1).Value = NextTicker) Then
                Sum = Sum + Cells(i, 7).Value
            Else
                Sum = Sum + Cells(i, 7).Value
                Cells(Result_row, 9).Value = Cells(i, 1).Value
                Cells(Result_row, 10).Value = Sum
                Result_row = Result_row + 1
                Sum = 0
            End If
        Next i
    Next ws
        
End Sub
' ==========================================================

