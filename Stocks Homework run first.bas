Attribute VB_Name = "Module1"
Sub stocks()

'Run first



'Define Last row of Raw Data


LastRow = Cells(Rows.Count, 1).End(xlUp).Row


'Headers
    
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    
'Variables

    Dim Ticker_Name As String
    Dim Ticker_Total As Double
    Dim Yearly_Change As Double
    Ticker_Total = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
       
'Start at very first open amount

     
   Year_Open = Cells(2, 3).Value
       
    
    For i = 2 To LastRow
     
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Ticker_Name = Cells(i, 1).Value
        
        Ticker_Total = Ticker_Total + Cells(i, 7).Value
        
        Year_close = Cells(i, 6).Value
         
            Yearly_Change = Year_close - Year_Open
            
            'If statement for Year_open to not equal zero
    If Year_Open = 0 Then
            
            Yearly_Percent_Change = "N/A"
            
        Else: Yearly_Percent_Change = (Yearly_Change / Year_Open)
            
    End If
    
    
            
     'Summary Table for output
     
       
        range("J" & Summary_Table_Row).Value = Ticker_Name
        
        range("M" & Summary_Table_Row).Value = Ticker_Total
        
        range("K" & Summary_Table_Row).Value = Yearly_Change
   
        range("L" & Summary_Table_Row).Value = Yearly_Percent_Change
        
        Summary_Table_Row = Summary_Table_Row + 1
                
 'After first loop then use first open amount of next ticker number
 
        Year_Open = Cells(i + 1, 3).Value
        
        Ticker_Total = 0
      
        
    Else: Ticker_Total = Ticker_Total + Cells(i, 7).Value
    
         
    End If
    
      
     Next i
     
     
     
'Define Last Row of Summary table

     
    LastRow = Cells(Rows.Count, 11).End(xlUp).Row
   
'Loop through summary table again to shade for value

    
    For j = 2 To LastRow
    
              
     
    If Cells(j, 11).Value >= 0 Then
        Cells(j, 11).Interior.ColorIndex = 4
        Cells(j, 12).NumberFormat = "0.00%"
        
    Else: Cells(j, 11).Interior.ColorIndex = 3
          Cells(j, 12).NumberFormat = "0.00%"
    
    
    
    End If
    
    
    
  
  Next j
  
  

  
  
  
   
End Sub


