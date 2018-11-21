Sub hw2()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        result_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ws.Range("I1").Value = " Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("I1:L1").Columns.AutoFit
        
        Sum = 0
        open_index = 2
        Result_row = 2
        
        For i = 2 To LastRow
            NextTicker = Cells(i + 1, 1).Value
            
            If (Cells(i, 1).Value = NextTicker) Then
                Sum = Sum + Cells(i, 7).Value
            Else
                Sum = Sum + Cells(i, 7).Value
                
                ' ticker
                Cells(Result_row, 9).Value = Cells(i, 1).Value
                
                'total volume
                Cells(Result_row, 10).Value = Sum
                
                'yearly change
                Cells(Result_row, 11).Value = Cells(i, 6).Value - Cells(open_index, 3).Value
                
                'Percent Change
                Cells(Result_row, 12).Value = Cells(Result_row, 11).Value / Cells(open_index, 3).Value
                Cells(Result_row, 12).NumberFormat = "0.00%"  'formating to %

                Result_row = Result_row + 1
                Sum = 0
                open_index = i + 1
                
            End If
        Next i
    Next ws
End Sub
' ==========================================================
