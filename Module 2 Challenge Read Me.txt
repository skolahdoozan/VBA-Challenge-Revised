Attribute VB_Name = "Module1"
Sub Challenge2():
Dim wb As Workbook
Dim ws As Worksheet
Dim NumberofSheets As Long
NumberofSheets = ActiveWorkbook.Worksheets.Count

For Each ws In Worksheets

' DEFINING HEADER AND REQUESTED CELLS
Dim rng As Range
Set rng = ws.Range("L1")
ws.Cells(1, 13) = "Open"
ws.Cells(1, 14) = "Close"
ws.Cells(1, 15) = "Difference"
ws.Cells(1, 16) = "% Change"
ws.Cells(1, 17) = "Total Volume for the Year (millions)"
ws.Cells(2, 20) = "Greatest % Increase"
ws.Cells(3, 20) = "Greatest % Decrease"
ws.Cells(4, 20) = "Greatest Volume (millions)"


' FINDS THE UNIQUE TICKER VALUES
ws.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=rng, Unique:=True

' FINDS THE NUMBER OF UNIQUE TICKERS
Dim EndRow As Long
EndRow = rng.End(xlDown).Row

' FINDS THE OPENING AND CLOSING VALUES FOR THE YEAR
Dim RowFirst As Long
Dim RowLast As Long

For i = 2 To EndRow
RowFirst = ws.Range("A:A").Find(ws.Cells(i, 12), lookat:=xlWhole, SearchDirection:=xlNext).Row
RowLast = ws.Range("A:A").Find(ws.Cells(i, 12), lookat:=xlWhole, SearchDirection:=xlPrevious).Row
ws.Cells(i, 13) = Format(ws.Cells(RowFirst, 3), "currency")
ws.Cells(i, 14) = Format(ws.Cells(RowLast, 6), "currency")
Next i

' FINDS THE CHANGE IN VALUES AS WELL AS THE CHANGE PERCENTAGE
Dim PercentChange As Double
For j = 2 To EndRow
ws.Cells(j, 15) = ws.Cells(j, 14) - ws.Cells(j, 13)
If ws.Cells(j, 15) < 0 Then
    ws.Cells(j, 15).Interior.Color = RGB(250, 0, 0)
Else
    ws.Cells(j, 15).Interior.Color = RGB(0, 250, 0)
End If
PercentChange = (ws.Cells(j, 15)) / (ws.Cells(j, 13))
ws.Cells(j, 16) = Format(PercentChange, "Percent")
Next j

' FINDS THE TOTAL VOLUME OF TRADED STOCK IN A YEAR
For k = 2 To EndRow
ws.Cells(k, 17) = Format(WorksheetFunction.SumIfs(ws.Range("H:H"), ws.Range("A:A"), ws.Cells(k, 12)), "Scientific")
Next k

' FINDS THE VALUE OF THE BEST AND WORST PERFORMANCE FOR TRADED STOCKS AS WELL AS THE MOST TRADED STOCK
Dim MaxPercentage As Variant
Dim MinPercentage As Variant
Dim MaxVolume As Long
MaxPercentage = Format(WorksheetFunction.Max(ws.Range("P:P")), "Percent")
MinPercentage = Format(WorksheetFunction.Min(ws.Range("P:P")), "Percent")
MaxVolume = Format(WorksheetFunction.Max(ws.Range("Q:Q")), "Scientific")
Dim Row1 As Long
Dim Row2 As Long
Dim Row3 As Long

' FINDS THE STOCK NAME FOR THE PREVIOUS DEFINED VALUES
ws.Cells(2, 22) = MaxPercentage
ws.Cells(3, 22) = MinPercentage
ws.Cells(4, 22) = MaxVolume
Row1 = ws.Columns(16).Find(MaxPercentage).Row
Row2 = ws.Columns(16).Find(MinPercentage).Row
Row3 = ws.Columns(17).Find(MaxVolume).Row

ws.Cells(2, 21) = ws.Cells(Row1, 12).Value
ws.Cells(3, 21) = ws.Cells(Row2, 12).Value
ws.Cells(4, 21) = ws.Cells(Row3, 12).Value

Next ws

     
End Sub


