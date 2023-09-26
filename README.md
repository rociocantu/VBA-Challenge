Sub stocks()
For Each ws In Worksheets
Dim ticker As String
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Price_Row As Long
Price_Row = 2
Total = 0
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
      
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To Lastrow:
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
Total = Total + ws.Range("G" & i).Value
ws.Range("I" & Summary_Table_Row).Value = ticker
ws.Range("L" & Summary_Table_Row).Value = Total
Open_Price = ws.Range("C" & Price_Row).Value
Close_Price = ws.Range("F" & i).Value
Yearly_Change = Close_Price - Open_Price
If Open_Price = 0 Then
    Percent_Change = 0
Else
    Percent_Change = Yearly_Change / Open_Price
End If
ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
ws.Range("K" & Summary_Table_Row).Value = Percent_Change
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
If ws.Range("J" & Summary_Table_Row).Value > 0 Then ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 Then
    Summary_Table_Row = Summary_Table_Row + 1
    Price_Row = i + 1
    Total = 0
Else
    Total = Total + ws.Range("G" & i).Value
    End If
Next i

Dim yearLastRow As Long
yearLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

For i = 2 To yearLastRow
If ws.Cells(i, 10).Value >= 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
Else
    ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i
    
Dim percentLastRow As Long
percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
Dim percent_max As Double
percent_max = 0
Dim percent_min As Double
percent_min = 0

For i = 2 To percentLastRow
If percent_max < ws.Cells(i, 11).Value Then
    percent_max = ws.Cells(i, 11).Value
    ws.Cells(2, 17).Value = percent_max
    ws.Cells(2, 17).Style = "Percent"
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
ElseIf percent_min > ws.Cells(i, 11).Value Then
    percent_min = ws.Cells(i, 11).Value
    ws.Cells(3, 17).Value = percent_min
    ws.Cells(3, 17).Style = "Percent"
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    End If
Next i

Dim totalVolumeRow As Long
totalVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
Dim totalVolumeMax As Double
totalVolumeMax = 0
For i = 2 To totalVolumeRow
    If totalVolumeMax < ws.Cells(i, 12).Value Then
    totalVolumeMax = ws.Cells(i, 12).Value
    ws.Cells(4, 17).Value = totalVolumeMax
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    End If
Next i
Next ws
End Sub
