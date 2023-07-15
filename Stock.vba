Sub Stock():

'dimensions
Dim ws As Worksheet
Dim total As Double
Dim openprice As Double
Dim closeprice As Double
Dim yearlychange As Double
Dim percentchange As Double
Dim totalvolume As Double
Dim Summaryrowcount As Integer
Dim ticker As String
Dim maxpercentchange As Double
Dim minpercentchange As Double
Dim maxtotalvolume As Double




For Each ws In Worksheets
total = 0
Summaryrowcount = 2
maxpercentchange = 0
minpercentchange = 0
maxtotalvolume = 0
 'row number of the last row
RowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

For i = 2 To RowCount
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
totalvolume = totalvolume + ws.Cells(i, 7).Value
closeprice = ws.Cells(i, 6).Value
yearlychange = closeprice - openprice
percentchange = yearlychange / openprice

ws.Cells(Summaryrowcount, 9).Value = ticker
ws.Cells(Summaryrowcount, 10).Value = yearlychange
ws.Cells(Summaryrowcount, 11).Value = percentchange
ws.Cells(Summaryrowcount, 12).Value = totalvolume

If percentchange > 0 Then
ws.Cells(Summaryrowcount, 10).Resize(1, 2).Interior.ColorIndex = 4
Else: ws.Cells(Summaryrowcount, 10).Resize(1, 2).Interior.ColorIndex = 3
End If


Summaryrowcount = Summaryrowcount + 1

If percentchange > maxpercentchange Then
maxpercentchange = percentchange
ws.Cells(2, 16).Value = maxpercentchange
ws.Cells(2, 17).Value = ticker
 End If
 
If percentchange < minpercentchange Then
minpercentchange = percentchange
ws.Cells(3, 16).Value = minpercentchange
ws.Cells(3, 17).Value = ticker
End If


If totalvolume > maxtotalvolume Then
maxtotalvolume = totalvolume
ws.Cells(4, 16).Value = maxtotalvolume
ws.Cells(4, 17).Value = ticker
End If





ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
openprice = ws.Cells(i, 3).Value
totalvolume = ws.Cells(i, 7).Value



Else
totalvolume = totalvolume + ws.Cells(i, 7).Value
End If



Next i
Next ws


End Sub