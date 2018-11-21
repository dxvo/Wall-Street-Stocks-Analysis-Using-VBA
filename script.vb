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
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Increase"
        ws.Range("N4").Value = "Greatest Total Volume"
        
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
                If (Cells(Result_row, 11).Value > 0) Then ' % change positive -> green
                    Cells(Result_row, 11).Interior.ColorIndex = 4
                ElseIf (Cells(Result_row, 11).Value < 0) Then
                    Cells(Result_row, 11).Interior.ColorIndex = 3
                End If
                
                'Percent Change
                If (Cells(open_index, 3).Value <> 0) Then ' if open value is not 0
                    Cells(Result_row, 12).Value = Cells(Result_row, 11).Value / Cells(open_index, 3).Value
                    Cells(Result_row, 12).NumberFormat = "0.00%"  'formating cells to %
                Else
                    Cells(Result_row, 12).Value = 0
                    Cells(Result_row, 12).NumberFormat = "0.00%"  'formating cells to %
                End If
            
                Result_row = Result_row + 1
                Sum = 0
                open_index = i + 1
                
            End If
        Next i
    
    result_last_row = ws.Cells(Rows.Count, 9).End(xlUp).Row 'last row of result
    
    greatest_increase = 0
    greatest_decrease = 0
    greatest_volume = 0
    increase_index = 0
    decrease_index = 0
    volume_index = 0
    
    
    For i = 2 To result_last_row
        If (Cells(i, 10).Value > greatest_volume) Then
            greatest_volume = Cells(i, 10).Value
            volume_index = i
        End If
        
        If (Cells(i, 12).Value > greatest_increase) Then
            greatest_increase = Cells(i, 12).Value
            increase_index = i
        ElseIf (Cells(i, 12).Value < greatest_decrease) Then
            greatest_decrease = Cells(i, 12).Value
            decrease_index = i
        End If
    Next i
    
    Range("O2").Value = Cells(increase_index, 9).Value
    Range("O3").Value = Cells(decrease_index, 9).Value
    Range("O4").Value = Cells(volume_index, 9).Value
    
    Range("P2").Value = Cells(increase_index, 12).Value
    Range("P2").NumberFormat = "0.00%"
    Range("P3").Value = Cells(decrease_index, 12).Value
    Range("P3").NumberFormat = "0.00%"
    Range("P4").Value = Cells(volume_index, 10).Value
    
    ws.Range("I1:P5").Columns.AutoFit
    Next ws
End Sub

'====================================================
'Comment:
    'Can also use min, min function for part3 of the exercise !