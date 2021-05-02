Attribute VB_Name = "Module1"
Sub stock_analysis():

'Create a script that will loop through all the stocks for one year and output the following information.

For Each ws In Worksheets

    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Volume As Double
    
    Dim StockOpen As Double
    Dim Stockclose As Double
    
    Dim lastrow As Double
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    ws.Range("i1") = "Ticker"
    ws.Range("j1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
 
Volume = 0

    Dim summary_table_row As Double
    summary_table_row = 2
    
    For i = 2 To lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        Ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
        
    ws.Range("I" & summary_table_row).Value = Ticker
    ws.Range("L" & summary_table_row).Value = Volume
    
    Volume = 0
    
    Stockclose = ws.Cells(i, 6)
    
    If StockOpen = 0 Then
        YearlyChange = 0
        PercentChange = 0
    
    Else:
        YearlyChange = Stockclose - StockOpen
        PercentChange = (Stockclose - StockOpen) / StockOpen
    End If
   
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.


ws.Range("J" & summary_table_row).Value = YearlyChange
ws.Range("K" & summary_table_row).Value = PercentChange
ws.Range("K" & summary_table_row).Style = "Percent"
ws.Range("K" & summary_table_row).NumberFormat = "0.00%"

summary_table_row = summary_table_row + 1

ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
    StockOpen = ws.Cells(i, 3)
    
Else: Volume = Volume + ws.Cells(i, 7).Value

End If

    Next i
    
For j = 2 To lastrow

'You should also have conditional formatting that will highlight positive change in green and negative change in red

If ws.Range("J" & j).Value > 0 Then
    ws.Range("J" & j).Interior.ColorIndex = 4
    
ElseIf ws.Range("J" & j).Value < 0 Then
        ws.Range("J" & j).Interior.ColorIndex = 3
        
End If
    Next j
        
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double
Dim GreatestVolume As Double

GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

For k = 2 To lastrow

    If ws.Cells(k, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(k, 11).Value
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").Style = "Percent"
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P2").Value = ws.Cells(k, 9).Value
        
    End If
    
    Next k

'Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
    
For l = 2 To lastrow

    If ws.Cells(l, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(l, 11).Value
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").Style = "Percent"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P3").Value = ws.Cells(l, 9).Value
        
    End If
    
    Next l
    
For m = 2 To lastrow

    If ws.Cells(m, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(m, 12).Value
        ws.Range("Q4").Value = GreatestVolume
        ws.Range("P4").Value = ws.Cells(m, 9).Value
        
    End If
    
    Next m
    
ws.Columns("A:Q").AutoFit
       
    
Next ws

End Sub


