Attribute VB_Name = "Module1"
'create script that loops through all the stocks for one year and outputs
'the ticker symbol
'yearly change from the opening price at the beginning of the given year to the closing price at the end of the year
'percent change from the opening price at the beginning of a given year to the closing price at the end of that year
'the total stock volume of the stock

Sub Stock_Market()


'define variables
Dim ws As Worksheet
Dim Ticker As String
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Stock_Volume As Double
Dim Summary_Table_Row As Integer

'loop through all sheets
For Each ws In Worksheets
'add column header for summary
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'keep track of the location for each ticker in the summary table
Summary_Table_Row = 2
Previous_i = 1
Stock_Volume = 0

'loop through all the ticker symbols using last row because each ws has a different count
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LastRow
'check if we are stll w/i the same ticker and if it is not then
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'set the ticker
Ticker = ws.Cells(i, 1).Value
Previous_i = Previous_i + 1
'add the open price/close price
Open_Price = ws.Cells(Previous_i, 3).Value
Close_Price = ws.Cells(i, 6).Value
'create a loop to get the requested measures
For j = Previous_i To i
'add the stock volume
Stock_Volume = Stock_Volume + ws.Cells(j, 7).Value
Next j
If Open_Price = 0 Then
Percent_Change = Close_Price
Else
Yearly_Change = Close_Price - Open_Price
Percent_Change = Yearly_Change / Open_Price
End If
'fill in the summary table
ws.Cells(Summary_Table_Row, 9).Value = Ticker
ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
ws.Cells(Summary_Table_Row, 12).Value = Stock_Volume
ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
ws.Cells(Summary_Table_Row, 10).NumberFormat = "0.00"
Summary_Table_Row = Summary_Table_Row + 1
Stock_Volume = 0
Yearly_Change = 0
Percent_Change = 0
Previous_i = i
End If

Next i

'add functionality to the script column O,P,Q
'look through column K for Percent Change
K_LastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
Increase = 0
Decrease = 0
Greatest = 0
For k = 3 To K_LastRow
Last_K = k - 1
Current_K = ws.Cells(k, 11).Value
Previous_K = ws.Cells(Last_K, 11).Value
volume = ws.Cells(k, 12).Value
PreviousVol = ws.Cells(Last_K, 12).Value
'Find the increase
If Increase > Current_K And Increase > Previous_K Then
Increase = Increase
ElseIf Current_K > Increase And Current_K > Previous_K Then
Increase = Current_K
increase_name = ws.Cells(k, 9).Value
ElseIf Previous_K > Increase And Previous_K > Current_K Then
Increase = Previous_K
increase_name = ws.Cells(Last_K, 9).Value
End If
'Find the decrease
If Decrease < Current_K And Decrease < Previous_K Then
Decrease = Decrease
ElseIf Current_K < Increase And Current_K < Previous_K Then
Decrease = Current_K
decrease_name = ws.Cells(k, 9).Value
ElseIf Previous_K < Increase And Previous_K < Current_K Then
Decrease = Previous_K
decrease_name = ws.Cells(Last_K, 9).Value
End If
'Find the greatest volume
If Greatest > volume And Greatest > PreviousVol Then
Greatest = Greatest
ElseIf volume > Greatest And volume > PreviousVol Then
Greatest = volume
greatest_name = ws.Cells(k, 9).Value
ElseIf PreviousVol > Greatest And Previous_Vol > volume Then
Greatest = PreviousVol
greatest_name = ws.Cells(Last_K, 9).Value
End If
Next k
'add column header for table
ws.Range("N1").Value = "Column Name"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker Name"
ws.Range("P1").Value = "Value"
'fill in column header for table
ws.Range("O2").Value = increase_name
ws.Range("O3").Value = decrease_name
ws.Range("O4").Value = greatest_name
ws.Range("P2").Value = Increase
ws.Range("P3").Value = Decrease
ws.Range("P4").Value = Greatest
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"
'conditional formatting for J(Yearly Change) & K (Percent Change)

J_LastRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
'K_LastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
For j = 2 To J_LastRow
If ws.Cells(j, 10) > 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 4
Else
ws.Cells(j, 10).Interior.ColorIndex = 3
End If
Next j

K_LastRow = ws.Cells(Rows.Count, "K").End(xlUp).Row
For j = 2 To K_LastRow
If ws.Cells(j, 10) > 0 Then
ws.Cells(j, 10).Interior.ColorIndex = 4
Else
ws.Cells(j, 10).Interior.ColorIndex = 3
End If
Next j
 

Next ws

End Sub

