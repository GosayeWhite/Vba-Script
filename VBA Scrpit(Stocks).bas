Attribute VB_Name = "Module1"
 Sub stockData()
 Dim ws As Worksheet
For Each ws In Worksheets
 ws.Activate
'Set headers
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Stock Total Volume"

 'Declare the data
 

Dim stock_name As String
Dim next_ticker As String
Dim previous_ticker As String
stock_total = 0
Dim summary_table As Double
summary_table = 2

Dim opening_price As Double
Dim closing_price As Double
Dim percent_decrease_ticker As String
Dim percent_increase_ticker As String
Dim last_row As Double

last_row = Range("A:A").SpecialCells(xlCellTypeLastCell).Row
         
For i = 2 To last_row

    stock_name = Cells(i, 1).Value
    next_ticker = Cells(i + 1, 1).Value
    previous_ticker = Cells(i - 1, 1).Value

    'stock volume total
    stock_total = stock_total + Cells(i, 7).Value
    'Yearly Change Cell K, Positive Integers
    PercentIncreaseNumber = Cells(2, "L").Value
    'Yearly Change Cell K, Negative Integers
    PercentDecreaseNumber = Cells(2, "L").Value
   If next_ticker <> stock_name Then
      
      closing_price = Cells(i, 6).Value
       Range("J" & summary_table).Value = stock_name
       Range("M" & summary_table).Value = stock_total
       stock_total = 0
       
       Range("K" & summary_table).Value = closing_price - opening_price
       If Range("K" & summary_table).Value > 0 Then
           Range("K" & summary_table).Interior.ColorIndex = 4
       Else
           Range("K" & summary_table).Interior.ColorIndex = 3
       End If
         
       
    Range("L" & summary_table).Value = FormatPercent((closing_price - opening_price) / opening_price, 2)
       If Range("L" & summary_table).Value > 0 Then
           Range("L" & summary_table).Interior.ColorIndex = 4
       Else
           Range("L" & summary_table).Interior.ColorIndex = 3
       End If

      If Range("L" & summary_table).Value > percent_increase_number Then
           percent_increase_number = Range("L" & summary_table).Value
           percent_increase_ticker = Range("J" & summary_table).Value
       End If
       If Range("L" & summary_table).Value < percent_decrease_number Then
           percent_decrease_number = Range("L" & summary_table).Value
           percent_decrease_ticker = Range("J" & summary_table).Value
       End If
       
       summary_table = summary_table + 1
       
  ElseIf previous_ticker <> stock_name Then
    opening_price = Cells(i, 3).Value
  
  
  End If
  
Next i

Cells(3, 16).Value = "Greatest%Increase"
Cells(4, 16).Value = "Greatest%Decrease"
Cells(5, 16).Value = " GreatestTotalVolume"
Cells(2, 17).Value = "Ticker"
Cells(2, 18).Value = "Value"


Cells(3, 18).Value = FormatPercent(Application.WorksheetFunction.Max(Range("L2:L" & last_row)))
Cells(4, 18).Value = FormatPercent(Application.WorksheetFunction.Min(Range("L2:L" & last_row)))
Cells(5, 18).Value = Application.WorksheetFunction.Max(Range("M2:M" & last_row))

Max = Application.WorksheetFunction.Max(Range("L2:L" & last_row))
Min = Application.WorksheetFunction.Min(Range("L2:L" & last_row))
Total = Application.WorksheetFunction.Max(Range("M2:M" & last_row))

greatest_increase = WorksheetFunction.Match(Max, Range("L:L"), 0)
greatest_decrease = WorksheetFunction.Match(Min, Range("L:L"), 0)
greatest_total = WorksheetFunction.Match(Total, Range("M:M"), 0)

Cells(3, 17).Value = Cells(greatest_increase, 10).Value
Cells(4, 17).Value = Cells(greatest_decrease, 10).Value
Cells(5, 17).Value = Cells(greatest_total, 10).Value


Next ws
End Sub





