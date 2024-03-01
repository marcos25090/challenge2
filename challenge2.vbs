Sub tickerloopsForEachSheet()
    Dim ws As Worksheet
   
    For Each ws In ThisWorkbook.Sheets
        
        Call tickerloops(ws)
    Next ws
End Sub

Sub tickerloops(ws As Worksheet)
    Dim tickername As String
    Dim tickervol As Double
    Dim summarytickerrow As Integer
    Dim openingprice As Double
    Dim closingprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim lastrow As Long
    Dim lastrowtable As Long

    
    ws.Cells(1, 9).Value = "ticker"
    ws.Cells(1, 12).Value = "total stock vol"
    ws.Cells(1, 11).Value = "percent change"
    ws.Cells(1, 10).Value = "yearly change"

    tickervol = 0
    summarytickerrow = 2
    openingprice = ws.Cells(2, 3).Value

    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            tickername = ws.Cells(i, 1).Value
            
            tickervol = tickervol + ws.Cells(i, 7).Value
            
            ws.Range("I" & summarytickerrow).Value = tickername
            
            ws.Range("L" & summarytickerrow).Value = tickervol
            
            closingprice = ws.Cells(i, 6).Value
            
            yearlychange = (closingprice - openingprice)
            
            ws.Range("J" & summarytickerrow).Value = yearlychange
            If openingprice = 0 Then
                percentchange = 0
            Else
                percentchange = yearlychange / openingprice
            End If
            
            ws.Range("K" & summarytickerrow).Value = percentchange
            
            ws.Range("K" & summarytickerrow).NumberFormat = "0.00%"
            
            summarytickerrow = summarytickerrow + 1
            
            tickervol = 0
            
            openingprice = ws.Cells(i + 1, 3)
        Else
            tickervol = tickervol + ws.Cells(i, 7).Value
        End If
    Next i

    lastrowtable = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

    For i = 2 To lastrowtable
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 10
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i

    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    For i = 2 To lastrowtable
        
        If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowtable)) Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowtable)) Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowtable)) Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        
        End If
    Next i
End Sub
