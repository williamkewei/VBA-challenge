Sub stock_analysis()
'Creating variables to hold ticker code, last row,yearly_change,percentage change,summary_row,opening_price,closing price
For Each ws In Worksheets
    Dim lastrow As Double
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    Dim ticker As String
    Dim total_volume As Double
    total_volume = 0
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim summary_row As Integer
    summary_row = 2
    Dim opening_price As Double
    Dim closing_price As Double
'print values
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "percentage Change"
    ws.Cells(1, 12) = "Total Volume"
    
    For i = 2 To lastrow
    'reset stock volume, redefine opening price for the next lot of data
    total_volume = total_volume + ws.Cells(i, 7)
    If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
        opening_price = ws.Cells(i, 3)
    ElseIf ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
    ticker = ws.Cells(i, 1)
    'define percentage change%,yearly_change,total volume
        closing_price = ws.Cells(i, 6)
        yearly_change = (closing_price - opening_price)
            If opening_price = 0 Then
            percentage_change = 0
        Else
            percentage_change = yearly_change / opening_price
        End If
        
            ws.Cells(summary_row, 9) = ticker
            ws.Cells(summary_row, 10) = yearly_change
            ws.Cells(summary_row, 11) = percentage_change
            ws.Cells(summary_row, 12) = total_volume
            total_volume = 0
            summary_row = summary_row + 1
            opening_price = 0
            clsoing_price = 0
        End If
    Next i
    
    ws.Range("K:K").NumberFormat = "0.00%"
    lastrow_new = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To lastrow_new
        If ws.Cells(i, 10) > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
'CHALLENGES

    Dim ticker1, ticker2, ticker3 As String
    Dim greatest_value As Double
    greatest_value = 0
    
    Dim lowest_value As Double
    lowest_value = 0
    
    Dim value As Double
    value = 0
    
    Dim summary_row2 As Integer
    summary_row2 = 2
    
    ws.Cells(1, 15) = "ticker"
    ws.Cells(1, 16) = "value"
    ws.Cells(2, 14) = "Greatest % Increase"
    ws.Cells(3, 14) = "Greatest % Decrease"
    ws.Cells(4, 14) = "Greatest Total Volume"
    
        For i = 2 To lastrow_new
            If ws.Cells(i, 11) > ws.Cells(i + 1, 11) And ws.Cells(i, 11) > greatest_value Then
            greatest_value = ws.Cells(i, 11)
            ticker1 = ws.Cells(i, 9)
            ElseIf ws.Cells(i, 11) < ws.Cells(i + 1, 11) And ws.Cells(i, 11) < lowest_value Then
                loest_value = ws.Cells(i, 11)
                ticker2 = ws.Cells(i, 9)
                ElseIf ws.Cells(i, 12) > ws.Cells(i + 1, 12) And ws.Cells(i, 12) > value Then
                    value = ws.Cells(i, 12)
                    ticker3 = ws.Cells(i, 9)
                    End If
                Next i
                
                ws.Cells(2, 15) = ticker1
                ws.Cells(3, 15) = ticker2
                ws.Cells(4, 15) = ticker3
                ws.Cells(2, 16) = greatest_value
                ws.Cells(2, 16).NumberFormat = "0.00%"
                ws.Cells(3, 16) = lowest_value
                ws.Cells(3, 16).NumberFormat = "0.00%"
                ws.Cells(4, 16) = value
                
            Next ws
            
        End Sub
        
                
            
    
    
End Sub
