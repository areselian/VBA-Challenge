Sub stockAnalysis():

Dim ticker As String
Dim openPrice As Double
Dim closePrice As Double
Dim totalStock As Double
Dim yearlyChange As Double
Dim greatInc As Double
Dim greatDec As Double
Dim greatTvol As Double
Dim Volume As Double
Volume = 0
Dim Row As Long
Row = 2
Dim Column As Integer
Column = 1
Dim i As Long
Dim z As Long

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    'Set Greatest % Inc and Dec, and Total Vol

    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"

    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"

    'set variables




    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    openPrice = Cells(2, Column + 2).Value

    'start the loop

    For i = 2 To lastRow
        
        Volume = Volume + ws.Cells(i, Column + 6).Value
        
        'set the initial numbers
        ' check if new ticker happemns
        If ws.Cells(i + 1, Column).Value <> ws.Cells(i, Column).Value Then
            'Check if no volume accrued
            If Volume = 0 Then
                ticker = ws.Cells(i, Column).Value
                ws.Range("J" & Row).Value = ticker
                ws.Range("K" & Row).Value = 0
                ws.Range("L" & Row).Value = 0
                
                'prepare for next loop
                Row = Row + 1
                yearlyChange = 0
                openPrice = ws.Cells(i + 1, Column + 2).Value
                
            Else
                ticker = ws.Cells(i, Column).Value
                ws.Cells(Row, Column + 8).Value = ticker
                closePrice = ws.Cells(i, Column + 5).Value
                yearlyChange = closePrice - openPrice
                ws.Cells(Row, Column + 9).Value = yearlyChange

            
                If openPrice = 0 Then
                    percentChange = 0

                Else
                    percentChange = yearlyChange / openPrice
                
                End If
            
                ws.Cells(Row, Column + 10).Value = percentChange
                ws.Cells(Row, Column + 10).NumberFormat = "0.00%"
                ws.Cells(Row, Column + 11).Value = Volume
                
                Row = Row + 1
                yearlyChange = 0
                Volume = 0
                openPrice = ws.Cells(i + 1, Column + 2).Value
            End If
        End If
    Next i
        
    YCLastRow = ws.Cells(Rows.Count, Column + 8).End(xlUp).Row
        
    For j = 2 To YCLastRow
        If (ws.Cells(j, Column + 9).Value > 0 Or ws.Cells(j, Column + 9).Value = 0) Then
                ws.Cells(j, Column + 9).Interior.ColorIndex = 10
        ElseIf ws.Cells(j, Column + 9).Value < 0 Then
                ws.Cells(j, Column + 9).Interior.ColorIndex = 3
        End If
    Next j
            
    For z = 2 To YCLastRow
    
    greatTvol = ws.Range("P2" & Row).Value
    
        If ws.Cells(z, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & YCLastRow)) Then
            ws.Cells(2, Column + 14).Value = ws.Cells(z, Column + 8).Value
            ws.Cells(2, Column + 15).Value = ws.Cells(z, Column + 10).Value
            ws.Cells(2, Column + 15).NumberFormat = "0.00%"
            
        ElseIf ws.Cells(z, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & YCLastRow)) Then
            ws.Cells(3, Column + 14).Value = ws.Cells(z, Column + 8).Value
            ws.Cells(3, Column + 15).Value = ws.Cells(z, Column + 10).Value
            ws.Cells(3, Column + 15).NumberFormat = "0.00%"
            
        ElseIf ws.Cells(z, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & YCLastRow)) Then
            ws.Cells(4, Column + 14).Value = ws.Cells(z, Column + 8).Value
            ws.Cells(4, Column + 15).Value = ws.Cells(z, Column + 11).Value
        
        End If
    
    Next z
Row = 2
Next ws

End Sub