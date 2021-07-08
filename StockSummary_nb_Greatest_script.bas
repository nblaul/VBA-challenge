Attribute VB_Name = "Module2"
Sub StockSummary_Greatest()

Dim Max As Double
Dim totalsummaryrows As Long
Dim ticker As String

For Each ws In ActiveWorkbook.Worksheets

GreatestIncrease = WorksheetFunction.Max(ws.Range("K:K"))
gi_ticker = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(GreatestIncrease, ws.Range("K:K"), 0))

GreatestDecrease = WorksheetFunction.Min(ws.Range("K:K"))
gd_ticker = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(GreatestDecrease, ws.Range("K:K"), 0))

GreatestVolume = WorksheetFunction.Max(ws.Range("L:L"))
gv_ticker = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(GreatestVolume, ws.Range("L:L"), 0))

ws.Cells(2, 14) = "Greatest % Increase"
ws.Cells(3, 14) = "Greatest % Decrease"
ws.Cells(4, 14) = "Greatest Volume"

ws.Cells(1, 15) = "Ticker"
ws.Cells(1, 16) = "Value"

ws.Cells(2, 16) = GreatestIncrease
ws.Cells(2, 16).NumberFormat = "0.00%"
ws.Cells(2, 15) = gi_ticker

ws.Cells(3, 15) = gd_ticker
ws.Cells(3, 16) = GreatestDecrease
ws.Cells(3, 16).NumberFormat = "0.00%"

ws.Cells(4, 15) = gv_ticker
ws.Cells(4, 16) = GreatestVolume

Next

End Sub
