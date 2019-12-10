Sub Color():
    Dim ws As Worksheet
    Dim i as Double

    For Each ws In Worksheets
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastRow 
            If (ws.Cells(i, 10).Value > 0 Or ws.Cells(i, 10).Value = 0) Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If 
        Next i
    Next ws
End Sub
