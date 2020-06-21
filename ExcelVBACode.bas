Attribute VB_Name = "Module1"
Sub hard()

Dim i As Long
Dim j As Integer
Dim lastrow As Long
Dim total As Double
Dim maxvol As Double
Dim maxinc As Double
Dim maxdec As Double
Dim indexmax As Long
Dim indexmaxdec As Long
Dim indexmaxinc As Long
Dim openprice As Double
Dim closeprice As Double
Dim ws As Worksheet


'Set ws = ActiveSheet

For Each ws In Worksheets
ws.Select

'initialize
j = 2
total = 0
openprice = 0
closeprice = 0

'find last row of data
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'initialize first index is 2
Cells(j, 9).Value = 2

'initialize
maxvol = Cells(2, 11).Value
maxinc = Cells(2, 13).Value
maxdec = Cells(2, 13).Value

'loop through all rows
For i = 2 To lastrow

'adding all the vol until changes in ticker occurs
total = total + Cells(i, 7).Value

'insert total volume into K
Cells(j, 11).Value = total


'if to detect ticker change
If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

'insert ticker names into J
Cells(j, 10).Value = Cells(i, 1)

'get open price and close price
openprice = Cells(Cells(j, 9).Value, 3).Value
closeprice = Cells(i, 6).Value

'insert price difference into L
Cells(j, 12).Value = closeprice - openprice

If openprice = 0 Then
Cells(j, 13).Value = 0
Else
'insert percent different into M
Cells(j, 13).Value = Cells(j, 12).Value / openprice
End If

'Change to percent format
Cells(j, 13).NumberFormat = "0.00%"

'increment index j
j = j + 1

'initial total volume after every dectected new ticker
total = 0

'insert index where new changes occur into I
Cells(j, 9).Value = i + 1

End If

Next i

lastj = j

'format the yearly change column
For j = 2 To lastj

If Cells(j, 12).Value >= 0 Then
Cells(j, 12).Interior.ColorIndex = 4
Else
Cells(j, 12).Interior.ColorIndex = 3
End If

Next j



'loop through all j
For j = 3 To lastj

'look for max total volume
If maxvol < Cells(j, 11).Value Then
maxvol = Cells(j, 11).Value
indexmax = j
End If

'look for greatest percent increase
If maxinc < Cells(j, 13).Value Then
maxinc = Cells(j, 13).Value
indexmaxinc = j
End If

'look for greatest percent decrease
If maxdec > Cells(j, 13).Value Then
maxdec = Cells(j, 13).Value
indexmaxdec = j
End If

Next j


'use application worksheet function max to check the results
'Cells(5, 17).Value = (Application.WorksheetFunction.Max(Range("K:K")))

'insert max volume value into R4
Cells(4, 18).Value = maxvol

'insert ticker that has that max value Q2
Cells(4, 17).Value = Cells(indexmax, 10).Value

'insert max increase value into R3
Cells(2, 18).Value = maxinc
Cells(2, 18).NumberFormat = "0.00%"

'insert ticker that has that max increase value Q3
Cells(2, 17).Value = Cells(indexmaxinc, 10).Value

'insert max decrese value into R4
Cells(3, 18).Value = maxdec
Cells(3, 18).NumberFormat = "0.00%"

'insert ticker that has that max decrease value Q4
Cells(3, 17).Value = Cells(indexmaxdec, 10).Value

'Labels Column Titles
Range("I1").Value = "Row where changes occur"

Range("J1").Value = "Ticker"
Range("K1").Value = "Total Stock Volume"
Range("L1").Value = "Yearly Change"
Range("M1").Value = "Percent Change"

Range("P2").Value = "Greatest % Increase"
Range("P3").Value = "Greatest % Decrease"
Range("P4").Value = "Greatest Total Volume"

Range("Q1").Value = "Ticker"
Range("R1").Value = "Value"

Next ws

End Sub





