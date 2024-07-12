Sub findstockSummary()
   ' This Module Needs to run after Module 5 as it Takes the Input from there
   
    Dim maxIncreasePct As Double
    Dim maxDecreasePct As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
 
    
    
    
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    Dim percentChange As Double
    Dim totalVolume As Double
    
     Dim ws As Worksheet
     Dim lastRow As Long
      
     
     
     maxIncreasePct = 0
     maxDecreasePct = 0
     maxVolume = 0
     
     For Each ws In ThisWorkbook.Worksheets
     
     
     
     
     lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
     MsgBox lastRow
     
 
     For x = 2 To lastRow
     
     'Input is I column for Ticker from Module 5
        Ticker = ws.Cells(x, 9).Value
        percentChange = ws.Cells(x, 11).Value
        totalVolume = ws.Cells(x, 12).Value
        
     ' Core logic
     
     If percentChange > maxIncreasePct Then
            maxIncreasePct = percentChange
            maxIncreaseTicker = Ticker
        ElseIf percentChange < maxDecreasePct Then
            maxDecreasePct = percentChange
            maxDecreaseTicker = Ticker
        End If
        
        If totalVolume > maxVolume Then
            maxVolume = totalVolume
            maxVolumeTicker = Ticker
        End If
        
    
    
    'Write Resultss
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = maxIncreaseTicker
     ws.Cells(2, 17).Value = maxIncreasePct
     
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = maxDecreaseTicker
    ws.Cells(3, 17).Value = maxDecreasePct
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = maxVolumeTicker
    ws.Cells(4, 17).Value = maxVolume
   
   Next x
Next ws
   
  End Sub
    
    
    
    
    
    
    
        
        
     
     
    
    
    

