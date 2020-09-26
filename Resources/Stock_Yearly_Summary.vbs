Sub yearlySummary()

Dim symbol As String
Dim openPrice As Double
Dim closePrice As Double
Dim yVolume As Double
Dim summaryRow As Integer
Dim ws As Worksheet
Dim i As Long
Dim maxIncrease As Double
Dim maxDecrease As Double
Dim maxVolume As Double

For Each ws In ThisWorkbook.Worksheets
ws.Activate

lastRow = CLng(Cells(Rows.Count, 1).End(xlUp).Row)

yVolume = 0

summaryRow = 1

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

For i = 2 To lastRow

    If Cells(i - 1, 1).Value <> Cells(i, 1) Then
        openPrice = Cells(i, 3).Value

    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1) Then
        
        symbol = Cells(i, 1).Value
        yVolume = yVolume + Cells(i, 7).Value
        summaryRow = summaryRow + 1
        closePrice = Cells(i, 6).Value
        Cells(summaryRow, 9).Value = symbol
        Cells(summaryRow, 10).Value = closePrice - openPrice
        Cells(summaryRow, 12).Value = yVolume
        yVolume = 0
            
            If openPrice = 0 Then
                Cells(summaryRow, 11).Value = 0
            Else
                Cells(summaryRow, 11).Value = (closePrice - openPrice) / openPrice
            End If
        
            If Cells(summaryRow, 10).Value > 0 Then
                Cells(summaryRow, 10).Interior.ColorIndex = 4
            ElseIf Cells(summaryRow, 10).Value < 0 Then
                Cells(summaryRow, 10).Interior.ColorIndex = 3
            End If
    
    Else
        
        yVolume = yVolume + Cells(i, 7).Value
    
    End If
    
        
Next i

maxIncrease = Cells(2, 11).Value
maxDecrease = Cells(2, 11).Value
maxVolume = Cells(2, 12).Value

For j = 2 To Cells(Rows.Count, 11).End(xlUp).Row
    If Cells(j + 1, 11).Value > maxIncrease Then
        maxIncrease = Cells(j + 1, 11).Value
        Cells(2, 16).Value = Cells(j + 1, 9).Value
        Cells(2, 17).Value = maxIncrease

    Else
    End If
    
    If Cells(j + 1, 11).Value < maxDecrease Then
        maxDecrease = Cells(j + 1, 11).Value
        Cells(3, 16).Value = Cells(j + 1, 9).Value
        Cells(3, 17).Value = maxDecrease

    Else
    End If
    
    If Cells(j + 1, 12).Value > maxVolume Then
        maxVolume = Cells(j + 1, 12).Value
        Cells(4, 16).Value = Cells(j + 1, 9).Value
        Cells(4, 17).Value = maxVolume

    Else
    End If

Next j

Next ws

End Sub
