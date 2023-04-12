Sub Bonus()

For Each ws In ThisWorkbook.Worksheets

Dim max_value As Double
Dim ticker_st As String
Dim min_value As Double
Dim max_vol As Variant

'Create headers in all worksheets
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"


'get last active row of summary table
Dim last_row_st As Long
last_row_st = ws.Cells(Rows.Count, 11).End(xlUp).Row
           
'Set values
max_value = 0
min_value = 0
max_vol = 0
      
'Loop through summary table
For i = 2 To last_row_st

If ws.Cells(i, 11).Value > max_value Then
max_value = ws.Cells(i, 11)
ticker_st = ws.Cells(i, 9)

'Print the max_value in the small summary table
ws.Range("Q2").Value = max_value
ws.Range("P2").Value = ticker_st

End If

If ws.Cells(i, 11).Value < min_value Then
min_value = ws.Cells(i, 11)
ticker_st = ws.Cells(i, 9)

'Print the min_value in the small summary table
ws.Range("Q3").Value = min_value
ws.Range("P3").Value = ticker_st

End If

If ws.Cells(i, 12).Value > max_vol Then
max_vol = ws.Cells(i, 12)
ticker_st = ws.Cells(i, 9)

'Print the max_vol in the small summary table
ws.Range("Q4").Value = max_vol
ws.Range("P4").Value = ticker_st

End If

Next i

ws.Range("Q2,Q3").NumberFormat = "0.00%"
ws.Columns("A:Q").AutoFit

      
      
Next ws


End Sub
