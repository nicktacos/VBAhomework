Sub TickerModule()
Dim ws As Worksheet
Dim ticker As String
Dim vol As Long
Dim yearopen As Double
Dim yearclose As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim SumTabRow As Integer

For Each ws In ThisWorkbook.Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    SumTabRow = 2
    
    For i = 2 To ws.UsedRange.Rows.Count
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        vol = ws.Cells(i, 7).Value
        yearopen = ws.Cells(i - 260, 3).Value
        yearclose = ws.Cells(i, 6).Value
        yearlychange = yearclose - yearopen
        percentchange = CDec(yearlychange / yearopen)
        
        ws.Cells(SumTabRow, 9).Value = ticker
        ws.Cells(SumTabRow, 10).Value = yearlychange
        ws.Cells(SumTabRow, 11).Value = percentchange
        ws.Cells(SumTabRow, 12).Value = vol
        SumTabRow = SumTabRow + 1
        vol = 0
        
        
        
        End If
    Next i
    
    For i = 2 To ws.UsedRange.Rows.Count
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
    ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    Next i
    
    
    
Next ws
End Sub

