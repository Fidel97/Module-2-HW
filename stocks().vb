Sub stocks()

 Dim ws As Worksheet
         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly  Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest%increase"
ws.Range("O3").Value = "Greatest% Decrease"
ws.Range("O4").Value = "Greatest TotalVolume"
  ' Set an initial variable for holding the Ticker name
  
  Dim Ticker As String
  
  
  Dim Total_volume As Double
  Total_volume = 0
  ' Keep track of the location for each credit card brand in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Dim opening_price_row As Integer
  opening_price_row = 2
  Dim yearly_change As Double
  yearly_change = 0
 RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
 opening_price = 0
  
  ' Loop through all stocks
  For i = 2 To RowCount
  
   'pull opening_price
   If opening_price = 0 Then
   opening_price = ws.Cells(i, 3).Value
   
   End If

    
    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      ' Set the Ticker name
      Ticker = ws.Cells(i, 1).Value
      ' Add to the Total_volume
      Total_volume = Total_volume + ws.Cells(i, 7).Value
      '
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      closin_price = ws.Cells(i, 6).Value
      yearly_change = closin_price - opening_price
      Percentage_change = yearly_change / opening_price
      ws.Range("J" & Summary_Table_Row).Value = yearly_change
      ws.Range("K" & Summary_Table_Row).Value = Percentage_change
      ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      ' Print the Total_volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_volume
      
      'Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      opening_price_row = opening_price_row + 1
      
      
      ' Reset the  Total_volume & yearly_change
      Total_volume = 0
      yearly_change = 0
      opening_price = 0
    
    ' If the cell immediately following a row is the same ticker..
    
    Else
      
      ' Add to the  Total_volume
       Total_volume = Total_volume + ws.Cells(i, 7).Value
    End If
  Next i
  
  
  For i = 2 To LastRow
            
               If ws.Range("J" & i).Value > 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 4
        
        ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
        
        Else
        ws.Range("J" & i).Interior.ColorIndex = 0
                End If
          
          Next i
     
  ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
  ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
  ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & RowCount))
  
  max_increase = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
  max_decrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
  max_volume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
  ws.Range("P2").Value = ws.Cells(max_increase + 1, 9)
  ws.Range("P3").Value = ws.Cells(max_decrease + 1, 9)
  ws.Range("P4").Value = ws.Cells(max_volume + 1, 9)
  
  Next ws
End Sub






