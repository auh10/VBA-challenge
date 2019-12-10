Sub SummaryTicker():
    Dim ws as Worksheet
    Dim wb as Workbook
    Dim ticker As String
    Dim yearChange As Double
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim percentChange As Double
    Dim totalVol As Double
    Dim summaryRow As Integer
    summaryRow = 2

    For Each ws In Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        yearOpen = ws.Cells(2, 3).Value

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Year Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest Increase %"
        ws.Cells(3, 14).Value = "Greatest Decrease %"
        ws.Cells(4, 14).Value = "Greatest Total Volume"

        For i = 2 To lastRow
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                ticker = ws.Cells(i, 1).Value
                ws.Cells(summaryRow, 9).Value = ticker
                yearClose = ws.Cells(i, 6).Value
                yearChange = yearClose - yearOpen
                ws.Cells(summaryRow, 10).Value = yearChange          
                If (yearOpen = 0 And yearClose = 0) Then
                    percentChange = 0
                ElseIf (yearOpen = 0 And yearClose <> 0) Then
                    percentChange = 1
                Else
                    percentChange = yearChange / yearOpen     
                    ws.Cells(summaryRow, 11).Value = percentChange
                    ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
                End If
                totalVol = totalVol + ws.Cells(i, 7).Value
                ws.Cells(summaryRow, 12).Value = totalVol
                summaryRow = summaryRow + 1
                totalVol = 0
            Else
                totalVol = totalVol + ws.Cells(i, 7).Value
            End If 
        Next i
    Next ws
End Sub

