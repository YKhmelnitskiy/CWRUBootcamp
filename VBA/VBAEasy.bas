Attribute VB_Name = "Module1"
Sub TickerTotalVolumeEasy()
Dim StockVolume As Variant
Dim LastRow As Long
StockVolume = 0
For Each ws In Worksheets
Dim spot As Integer
spot = 2
Dim WorksheetName As String
WorksheetName = ws.Name
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Total Stock Volume"
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   For i = 2 To LastRow
   If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
   StockVolume = StockVolume + ws.Cells(i, 7).Value
   Else
   StockVolume = StockVolume + ws.Cells(i, 7).Value
   ws.Cells(spot, 10).Value = StockVolume
   StockVolume = 0
   ws.Cells(spot, 9).Value = ws.Cells(i, 1).Value
   spot = spot + 1
   End If
   Next i
Next ws
End Sub
