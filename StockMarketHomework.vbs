Attribute VB_Name = "Module1"
Sub StockMarket()
For Each ws In Worksheets
  ' Set an initial variable
Dim Ticker As String
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "YearlyChange"
ws.Cells(1, 11).Value = "PercentageChange"
ws.Cells(1, 12).Value = "TotalStockVolume"
Dim TotalStockVolume As Double
Dim Lastrow As Long
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Dim Openprice As Double
Dim Closeprice As Double
Dim PercentageChange As Double
Dim YearlyChange As Double
Dim vMin, vMax As Double
Dim volMax As Double

' Loop through all rows
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
Openprice = ws.Cells(2, 3).Value
For i = 2 To Lastrow
' Check if we are still within value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    ' Set the Ticker value
    Ticker = ws.Cells(i, 1).Value
    ' Add to the Close price & yearly change
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    Closeprice = ws.Cells(i, 6).Value
    YearlyChange = Closeprice - Openprice
    ws.Range("J" & Summary_Table_Row).Value = YearlyChange
        If Openprice = 0 Then
             PercentChange = 0
        Else
             PercentageChange = (YearlyChange / Openprice)
        End If
    ws.Range("K" & Summary_Table_Row).Value = PercentageChange
    ' Print the values in the Summary Table
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume
    ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
    ' Add one to the summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    Openprice = ws.Cells(i + 1, 3).Value
    ' Reset the volume
    TotalStockVolume = 0
    Else
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    End If
Next i
Summary_Table_Rowlastrow = ws.Cells(Rows.Count, 10).End(xlUp).row
For j = 2 To Summary_Table_Rowlastrow
    If ws.Cells(j, 10).Value > 0 Then
         ws.Cells(j, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j

'Bonus

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "value"
ws.Cells(2, 14).Value = "Greatest % increase"
ws.Cells(3, 14).Value = "Greatest % decrease"
ws.Cells(4, 14).Value = "Greatest total volume"


vMin = Application.WorksheetFunction.Min(Columns("K"))
vMax = Application.WorksheetFunction.Max(Columns("K"))
volMax = Application.WorksheetFunction.Max(Columns("L"))

Range("P3") = (vMin)
Range("P2") = (vMax)
Range("P4") = (volMax)

Range("O2") = "=INDEX(I:I,match(P2,K:K,0))"
Range("O3") = "=INDEX(I:I,match(P3,K:K,0))"
Range("O4") = "=INDEX(I:I,match(P4,L:L,0))"

ws.Columns("N").AutoFit
ws.Columns("O").AutoFit
ws.Columns("P").AutoFit
ws.Columns("I").AutoFit
ws.Columns("J").AutoFit
ws.Columns("K").AutoFit
ws.Columns("L").AutoFit

Next ws

End Sub

        


