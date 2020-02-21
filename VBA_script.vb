Attribute VB_Name = "Module1"
Sub Stock():

Dim i As Long
Dim j As Long
Dim LastRow As Long
Dim Row_number As Long
Dim year_change As Double
Dim percentChange As Double
Dim open_RowNumber As Long
Dim total_stock_volume As Double
Dim sum As Double
Dim Max_number As Double
Dim Min_number As Double
Dim Max_totalVolume As Double

sum = 0

Row_number = 2
open_RowNumber = 2

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly change"
ws.Range("K1").Value = "Percent change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow

sum = sum + ws.Range("G" & i).Value
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'ticker symbol
ws.Range("I" & Row_number).Value = ws.Cells(i, 1).Value
'year change
year_change = ws.Range("F" & i).Value - ws.Range("C" & open_RowNumber).Value
ws.Range("J" & Row_number).Value = year_change
'percent change
    If ws.Range("C" & open_RowNumber).Value <> 0 Then
    percentChange = 100 * ws.Range("J" & Row_number).Value / ws.Range("C" & open_RowNumber).Value
    ws.Range("K" & Row_number).Value = percentChange & "%"
    End If

open_RowNumber = i + 1
ws.Range("L" & Row_number).Value = sum
Row_number = Row_number + 1
sum = 0

End If

Next i

For i = 2 To LastRow
If ws.Range("J" & i) > 0 Then

ws.Range("J" & i).Interior.ColorIndex = 4

ElseIf ws.Range("J" & i) < 0 Then

ws.Range("J" & i).Interior.ColorIndex = 3

End If
Next i

'Greatest increase
Max_number = ws.Range("K2").Value

For j = 2 To LastRow
If ws.Range("K" & j).Value > 0 Then

  If ws.Range("K" & j).Value > Max_number Then
    Max_number = ws.Range("K" & j).Value
    ws.Range("P2").Value = Max_number & "%"
    ws.Range("O2").Value = ws.Range("I" & j).Value
  End If

End If

Next j
ws.Range("P2").Value = 100 * ws.Range("P2").Value

'Greatest decrease
Min_number = ws.Range("K2").Value
For j = 2 To LastRow
If ws.Range("K" & j).Value < 0 Then

  If ws.Range("K" & j).Value < Min_number Then
    Min_number = ws.Range("K" & j).Value
    ws.Range("P3").Value = Min_number & "%"
    ws.Range("O3").Value = ws.Range("I" & j).Value
  End If

End If

Next j
ws.Range("P3").Value = 100 * ws.Range("P3").Value

'Greatest total volume with loop
Max_totalVolume = ws.Range("L2").Value

For j = 2 To LastRow

    If ws.Range("L" & j).Value > Max_totalVolume Then
     Max_totalVolume = ws.Range("L" & j).Value
     ws.Range("P4").Value = Max_totalVolume
     ws.Range("O4").Value = ws.Range("I" & j).Value
    End If

Next j


sum = 0
open_RowNumber = 2
Row_number = 2

Next ws

End Sub
