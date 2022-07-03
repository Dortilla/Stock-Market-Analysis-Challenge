Sub ResetSummary()

For Each ws In Worksheets

    ws.Range("I:I").Clear
    ws.Range("J:J").Clear
    ws.Range("K:K").Clear
    ws.Range("L:L").Clear
    ws.Range("N:N").Clear
    ws.Range("O:O").Clear
    ws.Range("P:P").Clear
Next ws

End Sub

