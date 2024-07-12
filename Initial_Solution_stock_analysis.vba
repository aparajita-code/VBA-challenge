Sub stock_quarterly_change_Q1():

' 1st attempt to solve the problem. Final solution is in other files

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim I As Long
    Dim Ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim volume As Double
    Dim Summary_Table_Row As Integer
    
    ' Setting the WS
     
     
     
    Set ws = ThisWorkbook.sheets("Q1")

    ' New summary table
    
    ws.range("K1").Value = "Percent Change"
    ws.range("L1").Value = "Total Stock Volume"
    ws.range("I1").Value = "Ticker"
    ws.range("J1").Value = "Quarterly Change"
    
' Summary table starts

    Summary_Table_Row = 2

    ' Looing through input data
    
    
    For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
             If year_open = 0 Then
                    year_open = ws.Cells(I, 3).Value
            End If
            
            If ws.Cells(I - 1, 1) = ws.Cells(I, 1) And ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
                    year_close = ws.Cells(I, 6).Value
                    yearly_change = year_close - year_open
                    Ticker = ws.Cells(I, 1).Value
                    
                    volume = volume + ws.Cells(I, 7).Value
                    
            
            

            ' Writing results to summary table
            
                      ws.range("J" & Summary_Table_Row).Value = yearly_change
                      ws.range("I" & Summary_Table_Row).Value = Ticker
                      ws.range("L" & Summary_Table_Row).Value = volume
                      volume = 0
                        
                        

            ' Core logic year_open is not zero)
            
                     If year_open <> 0 Then
                        ws.range("K" & Summary_Table_Row).Value = yearly_change / year_open
                        
                     End If

            ' Reset variables for next ticker
                    Summary_Table_Row = Summary_Table_Row + 1
                    volume = 0
                    year_open = 0
            Else
            
                   volume = volume + ws.Cells(I, 7).Value
                   'ws.Range("L" & Summary_Table_Row).Value = volume
                   
                    
         End If
        Next I
         
End Sub




Sub stock_quarterly_change_Q2():

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim I As Long
    Dim Ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim volume As Double
    Dim Summary_Table_Row As Integer
    
 
     
    Set ws = ThisWorkbook.sheets("Q2")

  
    ws.range("K1").Value = "Percent Change"
    ws.range("L1").Value = "Total Stock Volume"
    ws.range("I1").Value = "Ticker"
    ws.range("J1").Value = "Quarterly Change"
    

    Summary_Table_Row = 2

   
    For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
             If year_open = 0 Then
                    year_open = ws.Cells(I, 3).Value
            End If
            
            If ws.Cells(I - 1, 1) = ws.Cells(I, 1) And ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
                    year_close = ws.Cells(I, 6).Value
                    yearly_change = year_close - year_open
                    Ticker = ws.Cells(I, 1).Value
                    
                    volume = volume + ws.Cells(I, 7).Value
                    
            
            

       
                      ws.range("J" & Summary_Table_Row).Value = yearly_change
                      ws.range("I" & Summary_Table_Row).Value = Ticker
                      ws.range("L" & Summary_Table_Row).Value = volume
                      volume = 0
                        
                        

            
                     If year_open <> 0 Then
                        ws.range("K" & Summary_Table_Row).Value = yearly_change / year_open
                        
                     End If

          
                    Summary_Table_Row = Summary_Table_Row + 1
                    volume = 0
                    year_open = 0
            Else
            
                   volume = volume + ws.Cells(I, 7).Value
                 
                   
                    
         End If
        Next I
         
End Sub



Sub stock_quarterly_change_Q3():

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim I As Long
    Dim Ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim volume As Double
    Dim Summary_Table_Row As Integer
    
 
     
    Set ws = ThisWorkbook.sheets("Q3")

  
    ws.range("K1").Value = "Percent Change"
    ws.range("L1").Value = "Total Stock Volume"
    ws.range("I1").Value = "Ticker"
    ws.range("J1").Value = "Quarterly Change"
    

    Summary_Table_Row = 2

   
    For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
             If year_open = 0 Then
                    year_open = ws.Cells(I, 3).Value
            End If
            
            If ws.Cells(I - 1, 1) = ws.Cells(I, 1) And ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
                    year_close = ws.Cells(I, 6).Value
                    yearly_change = year_close - year_open
                    Ticker = ws.Cells(I, 1).Value
                    
                    volume = volume + ws.Cells(I, 7).Value
                    
            
            

       
                      ws.range("J" & Summary_Table_Row).Value = yearly_change
                      ws.range("I" & Summary_Table_Row).Value = Ticker
                      ws.range("L" & Summary_Table_Row).Value = volume
                      volume = 0
                        
                        

            
                     If year_open <> 0 Then
                        ws.range("K" & Summary_Table_Row).Value = yearly_change / year_open
                        
                     End If

          
                    Summary_Table_Row = Summary_Table_Row + 1
                    volume = 0
                    year_open = 0
            Else
            
                   volume = volume + ws.Cells(I, 7).Value
                 
                   
                    
         End If
        Next I
         
End Sub



Sub stock_quarterly_change_Q4():

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim I As Long
    Dim Ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim volume As Double
    Dim Summary_Table_Row As Integer
    
 
     
    Set ws = ThisWorkbook.sheets("Q4")

  
    ws.range("K1").Value = "Percent Change"
    ws.range("L1").Value = "Total Stock Volume"
    ws.range("I1").Value = "Ticker"
    ws.range("J1").Value = "Quarterly Change"
    

    Summary_Table_Row = 2

   
    For I = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
             If year_open = 0 Then
                    year_open = ws.Cells(I, 3).Value
            End If
            
            If ws.Cells(I - 1, 1) = ws.Cells(I, 1) And ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
                    year_close = ws.Cells(I, 6).Value
                    yearly_change = year_close - year_open
                    Ticker = ws.Cells(I, 1).Value
                    
                    volume = volume + ws.Cells(I, 7).Value
                    
            
            

       
                      ws.range("J" & Summary_Table_Row).Value = yearly_change
                      ws.range("I" & Summary_Table_Row).Value = Ticker
                      ws.range("L" & Summary_Table_Row).Value = volume
                      volume = 0
                        
                        

            
                     If year_open <> 0 Then
                        ws.range("K" & Summary_Table_Row).Value = yearly_change / year_open
                        
                     End If

          
                    Summary_Table_Row = Summary_Table_Row + 1
                    volume = 0
                    year_open = 0
            Else
            
                   volume = volume + ws.Cells(I, 7).Value
                 
                   
                    
         End If
        Next I
         
End Sub



Sub mainFunction():


stock_quarterly_change_Q1
stock_quarterly_change_Q2
stock_quarterly_change_Q3
stock_quarterly_change_Q4

End Sub

         
         
         
         
         
         
         
      
         
         
         
         

