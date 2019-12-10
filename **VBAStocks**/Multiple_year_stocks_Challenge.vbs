Sub Greatest():
    Dim maxIncrease as Double
    Dim ticker as String
    Dim summaryRow as Integer
    summaryRow = 2

    For Each ws in Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow
            If ws.Cells(i, "K").Value > ws.Cells(i + 1, "K").Value Then
            ticker = ws.Cells(i, 9).Value
            ws.Cells(summaryRow, 15).Value = ticker
            maxIncrease = Application.WorksheetFunction.Max(Column("K"))
            ws.Cells(summaryRow, 16).Value = maxIncrease
            ws.Cells(summaryRow, 16).NumberFormat = "0.00%"
            End If
        Next i
    Next ws 
End Sub
