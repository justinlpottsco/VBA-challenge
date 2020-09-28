Attribute VB_Name = "Module1"
Sub WallStreetStocks()
    
Dim Ticker As String
Dim lastrow As Double
Dim Total As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim i As Long

For Each ws In ThisWorkbook.Worksheets

'Add Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

Total = 0

'Set open price
Open_Price = ws.Cells(2, 3).Value

'Loop through column A tickers til end of column
For i = 2 To ws.Range("A1").CurrentRegion.End(xlDown).Row

'Sum up like tickers for total column G
If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
    Total = ws.Cells(i, 7).Value + Total
'List where last row differs from next row
lastrow = ws.Range("I1").CurrentRegion.Rows.Count + 1

'Place values in Ticker & Stock Volume columns...changes count based on lastrow
Else
    ws.Range("I" & lastrow).Value = ws.Cells(i, 1).Value
    totalA = ws.Cells(i, 7).Value
    Total = Total + totalA
    ws.Range("L" & lastrow).Value = Total
    Total = 0
   
Close_Price = ws.Cells(i, 6).Value

End If

'Calculate Yearly Change & Percent Change based on Open & Closed prices
Yearly_Change = Close_Price - Open_Price
    ws.Cells(2, 10).Value = Yearly_Change

If (Open_Price = 0 And Close_Price = 0) Then
    Percent_Change = 0
ElseIf (Open_Price = 0 And Close_Price <> 0) Then
    Percent_Change = -1
Else
    Percent_Change = Yearly_Change / Open_Price
    ws.Cells(2, 11).Value = Percent_Change
    ws.Cells(2, 11).NumberFormat = "0.00%"
End If

'Add conditional formatting
Select Case Change
    Case Is > 0
        ws.Range("J" & 2 + j).Interior.ColorIndex = 4
    Case Is < 0
        ws.Range("J" & 2 + j).Interior.ColorIndex = 3
End Select

Next i

Next ws

End Sub




