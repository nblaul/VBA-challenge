Attribute VB_Name = "Module1"
Sub StockSummary_nb()

Dim TotalRows As Long
Dim thisTicker As String
Dim nextTicker As String
Dim summaryRow As Long
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim TotalVolume As Double
Dim rowNum As Long
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets

summaryRow = 2
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"


TotalRows = ws.Cells(1, 1).End(xlDown).Row

For rowNum = 2 To TotalRows

    
    thisTicker = ws.Cells(rowNum, 1)
    nextTicker = ws.Cells(rowNum + 1, 1)
    PreviousTicker = ws.Cells(rowNum - 1, 1)
    TotalVolume = TotalVolume + ws.Cells(rowNum, 7)
    
    'find the opening price for the stock ticker

    If PreviousTicker <> thisTicker Then
        OpeningPrice = ws.Cells(rowNum, 3).Value
    
    End If
    
    'find the closing price for the stock ticker and summarize
    
    If thisTicker <> nextTicker Then
        ClosingPrice = ws.Cells(rowNum, 6).Value
        
        
    'add ticker to summary table
        ws.Cells(summaryRow, 9).Value = thisTicker
        
        
    'add PriceChange to summary table
    
        ws.Cells(summaryRow, 10).Value = ClosingPrice - OpeningPrice
        
        
    'add PercentChange to summary table
        If OpeningPrice = 0 Then
            ws.Cells(summaryRow, 11).Value = "Null"
        Else
            ws.Cells(summaryRow, 11).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
        'format PercentageChange as Percent
            ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
        End If
         'Format Cell Colors
            If ws.Cells(summaryRow, 10) < 0 Then
           ws.Cells(summaryRow, 10).Interior.ColorIndex = 3

            Else: ws.Cells(summaryRow, 10).Interior.ColorIndex = 4

            End If
    
    'find total Stock Volume to summary table
            'find total Stock Volume to summary table
        ws.Cells(summaryRow, 12).Value = TotalVolume
        summaryRow = summaryRow + 1
        TotalVolume = 0
        
    End If

Next rowNum


Next

End Sub
