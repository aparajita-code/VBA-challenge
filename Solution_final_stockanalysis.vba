Sub stock_quarterly_change_Q1():

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim I As Long
    Dim Ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim volume As Double
    Dim Summary_Table_Row As Integer
    
     

    
     
    For Each ws In ThisWorkbook.Worksheets

   
    
            ws.range("K1").Value = "Percent Change"
            ws.range("L1").Value = "Total Stock Volume"
            ws.range("I1").Value = "Ticker"
            ws.range("J1").Value = "Quarterly Change"
           
           ' Begining of the row
           
            Summary_Table_Row = 2

        ' Loop through rows of data
        
                For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
                        If year_open = 0 Then
                        year_open = ws.Cells(I, 3).Value
                End If
            
                        If ws.Cells(I - 1, 1) = ws.Cells(I, 1) And ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
                        year_close = ws.Cells(I, 6).Value
                        yearly_change = year_close - year_open
                        Ticker = ws.Cells(I, 1).Value
                        volume = volume + ws.Cells(I, 7).Value
                    
            
            

            ' summary table write
            
                      ws.range("J" & Summary_Table_Row).Value = yearly_change
                      ws.range("I" & Summary_Table_Row).Value = Ticker
                      ws.range("L" & Summary_Table_Row).Value = volume
                      volume = 0
                        
                        

            ' Core Logic % change
            
                     If year_open <> 0 Then
                        ws.range("K" & Summary_Table_Row).Value = yearly_change / year_open
                        ws.range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                        
                     End If

            ' Reset for next ticker
            
                    Summary_Table_Row = Summary_Table_Row + 1
                    volume = 0
                    year_open = 0
            Else
            
                   volume = volume + ws.Cells(I, 7).Value
                   'ws.Range("L" & Summary_Table_Row).Value = volume
                   
                    
         End If
        Next I
    Next ws
End Sub



         
         
         
         
         
         
         
      
         
         
         
         


