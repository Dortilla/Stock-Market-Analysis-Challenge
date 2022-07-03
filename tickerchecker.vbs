Sub ticker_checker()

Dim ws As Worksheet
Dim startingws As Worksheet
Sheets(1).Select
Set startingws = ActiveSheet

For Each ws In Sheets

ws.Activate


Dim TickerSymbol As String

Dim TSV As Double
TSV = 0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Dim OpenPrice As Variant
Dim ClosedPrice As Variant


    Dim YearlyChange As Double
    Dim PercentChange As Double
        Dim i As Double
        Dim FinalVariable As String
        i = 2
        FinalVariable = Cells(Rows.Count, 1).End(xlUp).Row
Range("I1").Value = "Ticker Symbol"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

    OpenPrice = Cells(i, 3).Value
        
For i = 2 To FinalVariable
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        TSV = TSV + Cells(i, 7).Value
    Else
    TickerSymbol = Cells(i, 1).Value
    Range("I" & Summary_Table_Row).Value = TickerSymbol
        TSV = TSV + Cells(i, 7).Value
        
            Cells(i + 1, 6).Select
            ActiveCell.Offset(-1, 0).Select
            ClosedPrice = ActiveCell.Value
            
            YearlyChange = ClosedPrice - OpenPrice
            If OpenPrice <> 0 Then
            PercentChange = YearlyChange / OpenPrice
            Else
            PercentChange = 0
            End If
             
             Cells(i + 1, 3).Select
             OpenPrice = ActiveCell.Value
             
             Range("J" & Summary_Table_Row).Value = YearlyChange
             Range("K" & Summary_Table_Row).Value = PercentChange
             Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
             Range("L" & Summary_Table_Row).Value = TSV
             Summary_Table_Row = Summary_Table_Row + 1
             TSV = 0
             TSV = TSV + Cells(i, 7).Value
    End If
Next i

For i = 2 To FinalVariable
            
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            ElseIf Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
            End If
Next i

Range("N2").Value = "Greatest % Increase:"
Range("N3").Value = "Greatest % Decrease:"
Range("N4").Value = "Greatest Total Volume:"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

Range("P2").Value = WorksheetFunction.Max(Range("K:K"))
Range("P2").NumberFormat = "0.00%"
Range("P3").Value = WorksheetFunction.Min(Range("K:K"))
Range("P3").NumberFormat = "0.00%"
Range("P4").Value = WorksheetFunction.Max(Range("L:L"))

For i = 2 To FinalVariable
    If Cells(i, 11).Value = Range("P2").Value Then
        Range("O2").Value = Cells(i, 9).Value
    End If
    
    If Cells(i, 11).Value = Range("P3").Value Then
        Range("O3").Value = Cells(i, 9).Value
    End If
    
    If Cells(i, 12).Value = Range("P4").Value Then
        Range("O4").Value = Cells(i, 9).Value
    End If
Next i

Cells(1, 9).EntireColumn.AutoFit
Cells(1, 10).EntireColumn.AutoFit
Cells(1, 11).EntireColumn.AutoFit
Cells(1, 12).EntireColumn.AutoFit
Cells(1, 14).EntireColumn.AutoFit
Cells(1, 15).EntireColumn.AutoFit
Cells(1, 16).EntireColumn.AutoFit




Next ws

starting_ws.Activate
    
End Sub

