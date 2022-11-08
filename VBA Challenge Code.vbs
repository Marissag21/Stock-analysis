Sub StockAnalysis()
    
      Dim ws As Worksheet

         ' Loop through all of the worksheets in the active workbook.
         For Each ws In Worksheets
         'Set titles to columns I through L
       ws.Cells(1, 9).Value = "Ticker"
      ws.Cells(1, 10).Value = "Yearly Change"
      ws.Cells(1, 11).Value = "Percent Change"
     ws.Cells(1, 12).Value = "Total Stock Volume"
    
 
    Dim RowCount As Long
    Dim i As Long
    Dim Ticker As String
    Dim YearlyChangeTotal As Double
    Dim LatestBlankRow As Integer
    LastestBlankRow = 0
    Dim Total As Double
    Dim OpenPrice As Double
    OpenPrice = 2
    
    Total = 0
    Dim PercentTotal As Double
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row


'Loop through rows
    For i = 2 To RowCount
       
       'If I look forward and see something different
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Add the volume
        Total = Total + ws.Cells(i, 7).Value
        'compute Yearly change and Percent change
        YearlyChangeTotal = ws.Cells(i, 6).Value - ws.Cells(OpenPrice, 3).Value
        PercentTotal = YearlyChangeTotal / ws.Cells(OpenPrice, 3).Value
        
        'Totals will reflect on columns I,J, and K
       ws.Range("I" & LatestBlankRow + 2).Value = ws.Cells(i, 1).Value
       ws.Range("J" & LatestBlankRow + 2).Value = YearlyChangeTotal
       ws.Range("K" & LatestBlankRow + 2).Value = PercentTotal
       
       'Change the format for columns J and K
       ws.Range("J" & LatestBlankRow + 2).NumberFormat = "0.00"
       
       ws.Range("K" & LatestBlankRow + 2).NumberFormat = "0.00%"

       ws.Range("L" & LatestBlankRow + 2).Value = Total
       
       'If the yearly change is positive, shade the cell green; if it is negative, shade the cell red
       If ws.Range("J" & LatestBlankRow + 2).Value > 0 Then
       ws.Range("J" & LatestBlankRow + 2).Interior.ColorIndex = 4
       ElseIf ws.Range("J" & LatestBlankRow + 2).Value < 0 Then
       ws.Range("J" & LatestBlankRow + 2).Interior.ColorIndex = 3
       
       End If
       
       
       Total = 0
       LatestBlankRow = LatestBlankRow + 1
       OpenPrice = i + 1
    Else
  Total = Total + ws.Cells(i, 7).Value
       


       'I'm done
    
    
 
 
 
     End If
    
    Next i
 
    Next ws
End Sub