Sub stockmarkethomework()
    Application.ScreenUpdating = False

        Dim stockOpen As Double
        Dim stockClose As Double
        Dim i As Long
        Dim h As Long
        Dim lastRow As Long

        For Each ws In Worksheets
    
        h = 2
        p = 1
        stockOpen = ws.Cells(2, 3)
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(h, 12).Value = 0
        
            For i = 2 To lastRow
                
                ws.Cells(h, 12).Value = ws.Cells(h, 12).Value + ws.Cells(i, 7).Value
                
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1) Then
                    stockClose = ws.Cells(i, 6)
                    ws.Cells(h, 9).Value = ws.Cells(i, 1).Value
                    ws.Cells(h, 10).Value = stockClose - stockOpen
                    ws.Cells(h, 9).Value = ws.Cells(i, 1).Value
                    ws.Cells(h, 12) = ws.Cells(h, 12).Value
                    
                    If ws.Cells(h, 12) <> 0 Then
                        ws.Cells(h, 11).Value = (stockClose - stockOpen) / stockOpen
                        ElseIf ws.Cells(h, 12) = 0 Then
                            ws.Cells(h, 11).Value = 0
                        End If
                
                    If ws.Cells(h, 11).Value > ws.Cells(2, 16).Value Then
                        ws.Cells(2, 15).Value = ws.Cells(i, 1).Value
                        ws.Cells(2, 16).Value = ws.Cells(h, 11).Value
                    End If
                        
                    If ws.Cells(h, 11).Value < ws.Cells(3, 16).Value Then
                        ws.Cells(3, 15).Value = ws.Cells(i, 1).Value
                        ws.Cells(3, 16).Value = ws.Cells(h, 11).Value
                    End If
                        
                    If ws.Cells(h, 12).Value > ws.Cells(4, 16).Value Then
                        ws.Cells(4, 15).Value = ws.Cells(i, 1).Value
                        ws.Cells(4, 16).Value = ws.Cells(h, 12).Value
                    End If
    
                    If ws.Cells(h, 10) > 0 Then
                        ws.Cells(h, 10).Interior.Color = RGB(169, 208, 142)
                    ElseIf ws.Cells(h, 10) < 0 Then
                        ws.Cells(h, 10).Interior.Color = RGB(255, 141, 147)
                    Else
                        ws.Cells(h, 10).Interior.Color = RGB(255, 217, 102)
                    End If
    
                    ws.Cells(h, 11).NumberFormat = "0.00%"
                    ws.Cells(h, 9).BorderAround ColorIndex:=1
                    ws.Cells(h, 10).BorderAround ColorIndex:=1
                    ws.Cells(h, 11).BorderAround ColorIndex:=1
                    ws.Cells(h, 12).BorderAround ColorIndex:=1
                    ws.Cells(h, 12).BorderAround ColorIndex:=1
                    
                    h = h + 1
                    ws.Cells(h, 12).Value = 0

                    If ws.Cells(i + 1, 3) <> 0 Then
                        stockOpen = ws.Cells(i + 1, 3)
                    Else
                        Do While stockOpen = 0
                            stockOpen = ws.Cells(i + p, 3).Value
                            p = p + 1
                        Loop
                    End If
                    
                End If
                
            Next i
            For o = 9 To 12
                ws.Cells(1, o).BorderAround ColorIndex:=1
            Next o
            
            For j = 14 To 16
                For k = 1 To 4
                    ws.Cells(k, j).BorderAround ColorIndex:=1
                Next k
            Next j
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 16) = ws.Cells(4, 16).Value * 100
        ws.Columns("I:L").AutoFit
        ws.Columns("N:P").AutoFit

        Next ws
        
    
    Application.ScreenUpdating = True
    End Sub
