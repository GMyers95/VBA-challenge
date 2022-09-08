Sub AlphabetSoupDeluxe()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Activate
    Dim ticker As String
    Dim turnout_table_row As Integer
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock As LongLong
    Dim starting_price As Double
    starting_price = Cells(2, 3).Value
    turnout_table_row = 2
    Cells(1, 10) = "Ticker"
    Cells(1, 11) = "Yearly Change"
    Cells(1, 12) = "Percent Change"
    Cells(1, 13) = "Total Stock Volume"
    For i = 2 To Cells(Rows.Count, 2).End(xlUp).Row
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            yearly_change = Cells(i, 6).Value - starting_price
            percent_change = (yearly_change / starting_price)
            total_stock = total_stock + Cells(i, 7).Value
            Range("j" & turnout_table_row).Value = ticker
            Range("k" & turnout_table_row).Value = yearly_change
            Range("l" & turnout_table_row).NumberFormat = "0.00%"
            Range("l" & turnout_table_row).Value = percent_change
            Range("m" & turnout_table_row).Value = total_stock
                With Range("k" & turnout_table_row).FormatConditions.Add(xlCellValue, xlGreater, "=0")
                .Interior.Color = vbGreen
               End With
            
               With Range("k" & turnout_table_row).FormatConditions.Add(xlCellValue, xlLess, "0")
                 .Interior.Color = vbRed
               End With
            turnout_table_row = turnout_table_row + 1
            total_stock = 0
            starting_price = Cells(i + 1, 3).Value

        Else
            total_stock = total_stock + Cells(i, 7).Value
        End If
    Next i
    For i = 2 To Cells(Rows.Count, 10).End(xlUp).Row
        Cells(2, 17) = "Greatest % Increase"
        Cells(3, 17) = "Greatest % Decrease"
        Cells(4, 17) = "Greatest Total Stock"
        Cells(1, 18) = "Ticker"
        Cells(1, 19) = "Value"
        Cells(2, 19).NumberFormat = "0.00%"
        Cells(3, 19).NumberFormat = "0.00%"
        Cells(2, 19).Value = WorksheetFunction.Max(Range("l2:l" & Cells(Rows.Count, 10).End(xlUp).Row))
            If Cells(i, 12).Value = Cells(2, 19).Value Then
            Cells(2, 18).Value = Cells(i, 10).Value
            End If
        Cells(3, 19).Value = WorksheetFunction.Min(Range("l2:l" & Cells(Rows.Count, 10).End(xlUp).Row))
            If Cells(i, 12).Value = Cells(3, 19).Value Then
            Cells(3, 18).Value = Cells(i, 10).Value
            End If
        Cells(4, 19).Value = WorksheetFunction.Max(Range("m2:m" & Cells(Rows.Count, 11).End(xlUp).Row))
            If Cells(i, 13).Value = Cells(4, 19).Value Then
            Cells(4, 18).Value = Cells(i, 10).Value
            End If
    Next i
    ws.Columns.AutoFit
    Next ws
End Sub


