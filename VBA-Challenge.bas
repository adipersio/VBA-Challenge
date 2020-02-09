Attribute VB_Name = "Module1"
Sub StockAnalysis()

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Volume As Long

Dim SummaryTableRow As Integer
SummaryTableRow = 2

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Volume"

For i = 2 To 70926

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    YearlyChange = (Cells(i, 6).Value - Cells(i, 3).Value)
    PercentChange = (YearlyChange / Cells(i, 3).Value) * 100
    Volume = Volume + Cells(i, 7).Value

    End If


    Range("K" & SummaryTableRow).NumberFormat = "0.00%"
     
    Range("I" & SummaryTableRow).Value = Ticker
    Range("J" & SummaryTableRow).Value = YearlyChange
    Range("K" & SummaryTableRow).Value = PercentChange
    Range("L" & SummaryTableRow).Value = Volume
    

    SummaryTableRow = SummaryTableRow + 1

    If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4

    If Cells(i, 10).Value <= 0 Then
    Cells(i, 10).Interior.ColorIndex = 3

    End If
    End If
End If

    Next i

Volume = 0

End Sub

