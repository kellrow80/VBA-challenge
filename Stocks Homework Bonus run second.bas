Attribute VB_Name = "Module2"
Sub bonusmaxmin()

'Run second



Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Volume"





Dim range_percent As range
Dim range_volume As range
Dim min_percent As Double
Dim max_percent As Double
Dim min_volume As Double
Dim max_volume As Double




    
    Set range_percent = range("L2:L20000")
    Set range_volume = range("M2:M20000")
    
    
    
    min_percent = Application.WorksheetFunction.min(range_percent)
    max_percent = Application.WorksheetFunction.max(range_percent)
    
    min_volume = Application.WorksheetFunction.min(range_volume)
    max_volume = Application.WorksheetFunction.max(range_volume)
    
    
    
    
    
    Cells(2, 17).Value = max_percent
    Cells(2, 17).NumberFormat = "0.00%"
    Cells(3, 17).Value = min_percent
    Cells(3, 17).NumberFormat = "0.00%"
    Cells(4, 17).Value = max_volume
    
    
    
    
  
    
    For i = 2 To 20000
    
    
    
    
    If Cells(i, 12).Value = max_percent Then
       Cells(2, 16).Value = Cells(i, 10).Value
       
       
    ElseIf Cells(i, 12).Value = min_percent Then
        Cells(3, 16).Value = Cells(i, 10)
        
    ElseIf Cells(i, 13).Value = max_volume Then
        Cells(4, 16).Value = Cells(i, 10)
        
      
        
    End If
    
    Next i
    
    
    
    
    
    
      


End Sub
