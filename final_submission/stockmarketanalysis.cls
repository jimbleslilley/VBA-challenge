Sub moderate()

    Dim ticker As String
    Dim stockOpen As Double
    Dim stockClose As Double
    Dim stockDiff As Double
    Dim stockCent As Double
    Dim stockVol As Long
    Dim lastRow As Long
    Dim i As Long
    Dim h As Long

    For Each ws In Worksheets

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    Cells(2, 16).Value = 0
    Cells(3, 16).Value = 0
    Cells(4, 16).Value = 0
    stockOpen = Cells(2, 3)
    h = 2
    lastRow = Cells(RoCount, 1).End(xlUp).Row

    For o = 9 To 12
        Cells(1, o).BorderAround ColorIndex:=1
    Next o
    
    For j = 14 To 16
        For k = 1 To 4
            Cells(k, j).BorderAround ColorIndex:=1
        Next k
    Next j
    
    For i = 2 To lastRow
    
        ticker = Cells(i, 1).Value
        stockClose = Cells(i, 6).Value
        stockVol = stockVol + Cells(i, 7).Value / 100
        
        If ticker <> Cells(i + 1, 1) Then
            
            
            Cells(h, 11).NumberFormat = "0.00%"
            
            
            Cells(h, 9).Value = ticker
            stockDiff = stockClose - stockOpen
            Cells(h, 10).Value = stockDiff
            Debug.Print (stockDiff)
            
            
            If stockOpen <> 0 Then
            stockCent = (stockClose - stockOpen) / stockOpen
            Else
            stockCent = 0
            End If
        
            Cells(h, 11) = stockCent
            Cells(h, 12) = stockVol
            Cells(h, 12) = Cells(h, 12) * 100
            
            If Cells(h, 10) > 0 Then
                Cells(h, 10).Interior.Color = RGB(169, 208, 142)
            ElseIf Cells(h, 10) < 0 Then
                Cells(h, 10).Interior.Color = RGB(255, 141, 147)
            Else
                Cells(h, 10).Interior.Color = RGB(255, 217, 102)
            End If
            
            Cells(h, 9).BorderAround ColorIndex:=1
            Cells(h, 10).BorderAround ColorIndex:=1
            Cells(h, 11).BorderAround ColorIndex:=1
            Cells(h, 12).BorderAround ColorIndex:=1
            Cells(h, 12).BorderAround ColorIndex:=1
            
            If stockCent > Cells(2, 16).Value Then
                Cells(2, 15).Value = ticker
                Cells(2, 16).Value = stockCent
            End If
            
            If stockCent < Cells(3, 16).Value Then
                Cells(3, 15).Value = ticker
                Cells(3, 16).Value = stockCent
            End If
            
            If stockVol > Cells(4, 16).Value Then
                Cells(4, 15).Value = ticker
                Cells(4, 16).Value = stockVol
                
            End If

            stockDiff = 0
            stockOpen = Cells(i + 1, 3).Value
            Debug.Print (stockClose)
            stockClose = 0
            stockCent = 0
            stockVol = 0
            h = h + 1

        End If

    Next i

    Cells(2, 16).NumberFormat = "0.00%"
    Cells(3, 16).NumberFormat = "0.00%"
    Cells(4, 16) = Cells(4, 16).Value * 100
    Columns("I:L").AutoFit
    Columns("N:P").AutoFit

    Next ws

End Sub