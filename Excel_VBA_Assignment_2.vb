
Sub VBA()
Dim ws As Worksheet
Dim Ticker As String
Dim Total_Stock_Volume As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Summary_Table_Row As Integer
Dim Maxy As Double
Dim Minp As Double
Dim Maxt As Double

For Each ws In Worksheets
ws.Cells(1, 10) = "Ticker"
ws.Cells(1, 11) = "Yearly_Change"
ws.Cells(1, 12) = "Percentage_Change"
ws.Cells(1, 13) = "Total_Stock_Volume"
ws.Cells(1, 17) = "Value"
ws.Cells(1, 16) = "Symbol"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(4, 15) = "Greatest Total Volume"
Summary_Table_Row = 2
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
Open_Price = ws.Cells(i, 3).Value
End If

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
Ticker = ws.Cells(i, 1).Value
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
Close_Price = ws.Cells(i, 6).Value
Yearly_Change = Close_Price - Open_Price
Percentage_Change = (Close_Price - Open_Price) / Open_Price
  
ws.Range("J" & Summary_Table_Row).Value = Ticker
ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
ws.Range("L" & Summary_Table_Row).Value = Percentage_Change
ws.Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
If ws.Range("K" & Summary_Table_Row) > 0 Then
  ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf ws.Range("K" & Summary_Table_Row) < 0 Then
  ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
Else
  ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 5
End If


Summary_Table_Row = Summary_Table_Row + 1
Total_Stock_Volume = 0
Yearly_Change = 0
Percentage_Change = 0

       
Else
Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value

End If
  
ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
  Next i

Set y = ws.Range("L2:L" & Summary_Table_Row)
        Maxy = Application.WorksheetFunction.Max(y)
ws.Cells(2, 17) = Maxy

increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row)), ws.Range("L2:L" & Summary_Table_Row), 0)
ws.Cells(2, 16) = ws.Cells(increase_number + 1, 10)

 Set p = ws.Range("L2:L" & Summary_Table_Row)
        Minp = Application.WorksheetFunction.Min(p)
ws.Cells(3, 17) = Minp

decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("L2:L" & Summary_Table_Row)), ws.Range("L2:L" & Summary_Table_Row), 0)
ws.Cells(3, 16) = ws.Cells(decrease_number + 1, 10)


Set t = ws.Range("M2:M" & Summary_Table_Row)
        Maxt = Application.WorksheetFunction.Max(t)
ws.Cells(4, 17) = Maxt
total_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("M2:M" & Summary_Table_Row)), ws.Range("M2:M" & Summary_Table_Row), 0)
ws.Cells(4, 16) = ws.Cells(total_number + 1, 10)

Next ws

End Sub