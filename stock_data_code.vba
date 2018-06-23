Sub stock_data()

    For Each ws In Worksheets

        Dim ticker_name As String
        Dim total_vol As Double
    
        total_vol = 0
    
        Dim summary_table_row As Integer
        summary_table_row = 2
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
        For i = 2 To lastRow:

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                ticker_name = Cells(i, 1).Value
                total_vol = total_vol + Cells(i, 7).Value
                
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Total Volume"

            Cells(summary_table_row, 9).Value = ticker_name
            Cells(summary_table_row, 10).Value = total_vol

            summary_table_row = summary_table_row + 1

            total_vol = 0

            Else

                total_vol = total_vol + Cells(i, 7).Value
            
            End If
        
    Next i
    
    Next ws 

End Sub
