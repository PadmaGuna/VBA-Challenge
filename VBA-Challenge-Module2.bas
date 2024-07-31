Attribute VB_Name = "Module12"
Sub qtr_anlysis()
  Dim ws As Worksheet
  Dim month1, month2 As Integer
  Dim qtr_open_found As Boolean
  Dim qtr_close_found As Boolean
  Dim Qtr_row_num As Integer
  
  For Each ws In ThisWorkbook.Worksheets
    no_of_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    
    ' Set the heading for Qtr Analysis
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Quarterly Change"
  ws.Cells(1, 11) = "Percent Change"
  ws.Cells(1, 12) = "Total Stock Volume"
  ws.Cells(1, 15) = "Ticker"
  ws.Cells(1, 16) = "Value"
  
  ws.Cells(2, 14) = "Greatest % Increase"
  ws.Cells(3, 14) = "Greatest % Decrease"
  ws.Cells(4, 14) = "Greatest Total Volume"
  
  ws.Columns("N").ColumnWidth = 25
  ws.Columns("L").ColumnWidth = 25
  
  Qtr_row_num = 1
    
  For i = 2 To no_of_rows
     ''Get ticker symbol for current row and next row
     ticker_sym1 = ws.Cells(i, 1)
     ticker_sym2 = ws.Cells(i + 1, 1)
     If ticker_sym1 = ticker_sym2 Then
         If qtr_open_found = False Then
            qtr_open_bid = ws.Cells(i, 3)
            qtr_open_found = True
            qtr_volume = ws.Cells(i, 7)
            Qtr_row_num = Qtr_row_num + 1
         Else
            qtr_volume = qtr_volume + ws.Cells(i, 7)
         End If
     Else
        ''Get last row values
        qtr_close_bid = ws.Cells(i, 6)
        qtr_close_found = True
        qtr_volume = qtr_volume + ws.Cells(i, 7)
            
        'write values of Quater Anaysis
        ws.Cells(Qtr_row_num, 9) = ticker_sym1
        ws.Cells(Qtr_row_num, 10) = qtr_close_bid - qtr_open_bid
        ''Change color of the cell
        If ws.Cells(Qtr_row_num, 10) <= 0 Then
            ' Set the Cell Colors to Red
            ws.Cells(Qtr_row_num, 10).Interior.ColorIndex = 3
        Else
            ' Set the Cell Colors to Green
            ws.Cells(Qtr_row_num, 10).Interior.ColorIndex = 4
        End If
        
        qtrly_pct_change = (ws.Cells(Qtr_row_num, 10) / qtr_open_bid)
        ws.Cells(Qtr_row_num, 11) = qtrly_pct_change
        ws.Range("K" & Qtr_row_num).NumberFormat = "0.00%"
        ws.Cells(Qtr_row_num, 12) = qtr_volume
        ws.Range("L" & Qtr_row_num).NumberFormat = "#,###"
    ''Get greatest %increase, decrease and volume
        If qtrly_pct_change < grt_pct_decrease Then
            grt_pct_decrease = qtrly_pct_change
            ws.Cells(3, 15) = ticker_sym1
        End If
        
        If qtrly_pct_change > grt_pct_increase Then
            grt_pct_increase = qtrly_pct_change
            ws.Cells(2, 15) = ticker_sym1
        End If
        
        If qtr_volume > grt_qtr_Volume Then
            grt_qtr_Volume = qtr_volume
            ws.Cells(4, 15) = ticker_sym1
        End If
     
        qtr_open_found = False
        qtr_close_found = False
        qtr_volume = 0
     
     End If
     
   Next i
   
   
   ws.Cells(2, 16) = grt_pct_increase
   ws.Range("P2").NumberFormat = "0.00%"
   
   ws.Cells(3, 16) = grt_pct_decrease
   ws.Range("P3").NumberFormat = "0.00%"
   
   ws.Cells(4, 16) = grt_qtr_Volume
   ws.Range("P4").NumberFormat = "#,###"
   
   grt_pct_increase = 0
   grt_qtr_Volume = 0
   grt_pct_decrease = 0
   
Next ws
End Sub

