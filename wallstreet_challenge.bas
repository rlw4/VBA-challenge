Attribute VB_Name = "Module1"
Sub wallstreet_challenge()

Dim open_val As Double
Dim close_val As Double

Dim ticker As String
Dim year_change As Double
Dim per_change As Double
Dim Total As Double

Dim row_placement As Integer
Dim num_rows As Double
Dim count_tick As Integer

ticker = Cells(2, 1).Value
open_val = Cells(2, 3).Value
row_placement = 2
num_rows = Cells(Rows.Count, 1).End(xlUp).Row
count_tick = 1

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"


For i = 2 To num_rows
    Total = Total + Cells(i, 7).Value
    
    If StrComp(Cells(i, 1).Value, Cells(i + 1, 1).Value) <> 0 Then
    
    close_val = Cells(i, 6)
    
    Cells(row_placement, 9).Value = ticker
    ticker = Cells(i + 1, 1).Value
    count_tick = count_tick + 1
    
    year_change = close_val - open_val
    Cells(row_placement, 10).Value = Round(year_change, 2)
    
    per_change = year_change / open_val
    Cells(row_placement, 11).Value = Round(per_change, 2)
    
    Cells(row_placement, 12).Value = Round(Total, 0)
    Total = 0
    
    row_placement = row_placement + 1
    
    End If
    
    Next i
    
For i = 2 To count_tick + 1
    If Cells(i, 10).Value < 0 Then
        Cells(i, 10).Interior.ColorIndex = 3
        
    ElseIf Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
        
    End If
    
   Next i
    
End Sub
