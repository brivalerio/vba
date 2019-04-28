    Sub stockMarket()

    For Each ws In Worksheets

    Dim ticker As String
    Dim currentTicker As String
    Dim nextTicker As String
    Dim total As Double
    Dim rowLimit As Double
    Dim summaryTableRow As Integer
    summaryTableRow = 2

    rowLimit = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
        
    For i = 2 To rowLimit
        currentTicker = ws.Cells(i, 1).Value
        nextTicker = ws.Cells(i + 1, 1).Value

            If currentTicker <> nextTicker Then
                ticker = currentTicker
                currentTicker = nextTicker

                total = total + ws.Cells(i, 7).Value

                ws.Range("I" & summaryTableRow).Value = ticker
                ws.Range("J" & summaryTableRow).Value = total

                total = 0
                summaryTableRow = summaryTableRow + 1
            
            Else
                total = total + ws.Cells(i, 7).Value
            
            End If

    Next i

    ws.Range("I1:J1").EntireColumn.AutoFit

    Next ws

    End Sub
