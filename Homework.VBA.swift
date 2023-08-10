Sub test()
For Each ws In Worksheets
Dim TickerCount As Double
Dim x As Double
TickerCount = 2
Dim y As Double
Dim percentchange As Double
y = 2
RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"
For x = 2 To RowCount
    If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
        ws.Cells(TickerCount, 9).Value = ws.Cells(x, 1).Value
        ws.Cells(TickerCount, 10).Value = ws.Cells(x, 6).Value - ws.Cells(y, 3).Value
            If ws.Cells(TickerCount, 10).Value < 0 Then
            ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
        End If
        If ws.Cells(y, 3).Value <> 0 Then
        ws.Cells(TickerCount, 11) = (ws.Cells(x, 6).Value - ws.Cells(y, 3).Value) / ws.Cells(y, 3).Value
        Else
        ws.Cells(TickerCount, 11).Value = 0
        End If
       
        

        ws.Cells(TickerCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(y, 7), ws.Cells(x, 7)))
        TickerCount = TickerCount + 1
    End If
Next x
Range("P2") = "%" & WorksheetFunction.Max(Range("K2:K" & RowCount)) * 100
Range("P3") = "%" & WorksheetFunction.Min(Range("K2:K" & RowCount)) * 100
Range("P4") = WorksheetFunction.Max(Range("L2:L" & RowCount))
    ' returns one less because header row not a factor
increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & RowCount)), Range("K2:K" & RowCount), 0)
volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & RowCount)), Range("L2:L" & RowCount), 0)
    ' final ticker symbol for  total, greatest % of increase and decrease, and average
Range("O2") = Cells(increase_number + 1, 9)
Range("O3") = Cells(decrease_number + 1, 9)
Range("O4") = Cells(volume_number + 1, 9)
Next ws
End Sub
