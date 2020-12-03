Attribute VB_Name = "Module1"
Sub StockMarket()

For Each ws In Worksheets
Dim WorksheetName As String
Dim TickerName As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim ChangeInPrice As Double
Dim ChangeInPercent As Double
Dim PercentChange As Double
Dim TotalTickerVolume As Double
Dim Summary_Table_Row As Integer
Dim LasrRow As Long

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
OpenPrice = ws.Cells(2, 3).Value
Summary_Table_Row = 2
TotalTickerVolume = 0
For i = 2 To LastRow
    If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
    TickerName = ws.Cells(i, 1).Value
    ClosePrice = ws.Cells(i, 6).Value
    ChangeInPrice = ClosePrice - OpenPrice
        If OpenPrice <> 0 Then
        ChangeInPercent = (ChangeInPrice / OpenPrice) * 100
        End If
    ws.Range("I" & Summary_Table_Row).Value = TickerName
    ws.Range("J" & Summary_Table_Row).Value = ChangeInPrice
    TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
        If ChangeInPrice > 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf ChangeInPrice <= 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
    ws.Range("K" & Summary_Table_Row).Value = (CStr(ChangeInPercent) & "%")
    ws.Range("L" & Summary_Table_Row).Value = TotalTickerVolume
    OpenPrice = ws.Cells(i + 1, 3).Value
    
    Summary_Table_Row = Summary_Table_Row + 1
    ChangeInPrice = 0
    ClosePrice = 0
    Else
        TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
    End If
Next i
Next ws
End Sub
