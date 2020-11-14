# vba_homework
week 2 homework

Sub StockTickers()

For Each ws In Worksheets

Dim WorksheetName As String
Dim Ticker As String
Dim rowcount As Integer
rowcount = 2
Dim i As Double
Dim Vol_Total As Double
Dim Yearly_Change As Double
Dim open_val As Integer
Dim rng As Range
Dim condition1 As FormatCondition, condition2 As FormatCondition






lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To lastrow
open_val = 2
Vol_Total = Vol_Total + Cells(i, 7).Value
Set rng = ws.Range("K" & rowcount)

WorksheetName = ws.Name


If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value
Vol_Total = Vol_Total + Cells(i, 7).Value
Yearly_Change = Cells(i, 6) - Cells(open_val, 3)
Percentage_Change = Round((Yearly_Change / Cells(open_val, 3) * 100), 3)
ws.Range("I1").Value = "Ticker"
ws.Range("j1").Value = "Vol"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percentage Change"
rng.FormatConditions.Delete




ws.Range("I" & rowcount) = Ticker
ws.Range("J" & rowcount) = Vol_Total
ws.Range("K" & rowcount) = Yearly_Change
ws.Range("L" & rowcount) = Percentage_Change & "%"
Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "0")
Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "0")
With condition1
    .Interior.Color = vbGreen
   End With

   With condition2
     .Interior.Color = vbRed
   End With

            
rowcount = rowcount + 1
Vol_Total = 0
open_value = i + 1
End If
Next i
Next ws




 
End Sub
