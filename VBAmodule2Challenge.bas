Attribute VB_Name = "Module1"
Sub VBA_challenge()
Dim Ticker As String
Dim j As Long
Dim y As Long
Dim Volume As Double
Dim LastRow As Double
Volume = 0
y = 2
StockNumber = 1

For Each ws In ThisWorkbook.Worksheets

j = 2

LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

ws.Range("I" & 1).Value = "ticker"
ws.Range("J" & 1).Value = "yearly change"
ws.Range("K" & 1).Value = "percentage change"
ws.Range("L" & 1).Value = "total stock volume"
For x = 2 To LastRow:
Volume = Volume + ws.Cells(x, 7).Value
    If ws.Cells(x, 1).Value <> ws.Cells(x + 1, 1).Value Then
        ws.Range("L" & j).Value = Volume
        Volume = 0
        Ticker = ws.Cells(x, 1).Value
        ws.Range("I" & j).Value = Ticker
        ClosingPrice = ws.Cells(x, 6).Value
        OpeningPrice = ws.Cells(y, 3).Value
        ws.Range("J" & j).Value = ClosingPrice - OpeningPrice
        ws.Range("K" & j).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
        ws.Range("K" & j).NumberFormat = "0.00%"
        Debug.Print Ticker
        Debug.Print ws.Name
        Debug.Print j
        If ws.Cells(j, 10).Value >= 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
            ws.Cells(j, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 3
            ws.Cells(j, 11).Interior.ColorIndex = 3
        End If
        StockNumber = StockNumber + 1
        y = x + 1
        j = j + 1
    End If
Next x
ws.Range("P" & 1).Value = "Ticker"
ws.Range("Q" & 1).Value = "Value"
ws.Range("O" & 2).Value = "Greatest % Increase"
ws.Range("O" & 3).Value = "Greatest % Decrease"
ws.Range("O" & 4).Value = "Greatest Total Volume"
Dim GI As Double
Dim GD As Double
Dim GV As Double
Dim IncN As String
Dim DCN As String
Dim Vol As String
GI = ws.Cells(2, 11).Value
GD = ws.Cells(2, 11).Value
GV = ws.Cells(2, 12).Value
IncN = ws.Cells(2, 9).Value
DCN = ws.Cells(2, 9).Value
Vol = ws.Cells(2, 9).Value
For i = 2 To StockNumber
    If ws.Cells(i, 11).Value > GI Then
        GI = ws.Cells(i, 11).Value
        IncN = ws.Cells(i, 9).Value
    End If
    If ws.Cells(i, 11).Value < GD Then
        GD = ws.Cells(i, 11).Value
        DCN = ws.Cells(i, 9).Value
    End If
    If ws.Cells(i, 12).Value > GV Then
        GV = ws.Cells(i, 12).Value
        Vol = ws.Cells(i, 9).Value
    End If
    Next i
    ws.Range("P2").Value = IncN
    ws.Range("P3").Value = DCN
    ws.Range("P4").Value = Vol
    ws.Range("Q2").Value = GI
    ws.Range("Q3").Value = GD
    ws.Range("Q4").Value = GV
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
Next ws
End Sub
