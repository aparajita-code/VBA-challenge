Sub Highlightcolors()


    Dim qtchg As range
    Dim rng As range
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
    Set rng = ws.range("J2:J1501")
    
    For Each qtchg In rng.Cells
        If qtchg.Value < 0 Then
        
        ' Red Cell color for negative values
        
            qtchg.Interior.Color = RGB(255, 0, 0)
            
        ElseIf qtchg.Value >= 0 Then
        ' Green celll color for negative values
            qtchg.Interior.Color = RGB(0, 255, 0)
            
        End If
    Next qtchg
    Next ws
       
End Sub
