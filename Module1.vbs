Sub Stocks()

Dim ws As Worksheet

'Create headers in all worksheets
For Each ws In ThisWorkbook.Worksheets

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"

'Variable for ticker_symbol
Dim ticker_symbol As String

'Variable for total stock volume
Dim total_stock_volume As Variant
total_stock_volume = 0

'Variable for Summary_table row
Dim summary_table_row As Integer
summary_table_row = 2

'Variable for opening_value
Dim opening_value As Double

'Variable for closing_value
Dim closing_value As Double

'Variable for yearly_change
Dim yearly_change As Double

'get last active row
Dim last_row As Long
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Variable for percent change
Dim percent_change As Double

'Loop through all stock lines
For i = 2 To last_row

 'Check for same ticker, if  not
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'Set ticker symbol
      ticker_symbol = ws.Cells(i, 1).Value
      
    'Set closing_value
      closing_value = ws.Cells(i, 6).Value
           
     'Set opening_price
     opening_value = ws.Cells(i - Application.WorksheetFunction.CountIf(ws.Range("A2:A" & i), ticker_symbol) + 1, 3).Value
           
     'Get yearly_change value
      yearly_change = closing_value - opening_value
      
      'Get percentage_change value
      percent_change = yearly_change / opening_value
           
    'Add to total_stock_volume
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
      
    'Print the ticker in the Summary Table
      ws.Range("I" & summary_table_row).Value = ticker_symbol
      
    'Print the total_stock_volume to the Summary Table
      ws.Range("L" & summary_table_row).Value = total_stock_volume
      
    'Print the yearly_change in the summary table
      ws.Range("J" & summary_table_row).Value = yearly_change
      
    'Print the percent_change in the summary table
      ws.Range("K" & summary_table_row).Value = percent_change
           
        'Conditional formatting green is 4 and red is 3
        If yearly_change > 0 Then
        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
        
        ElseIf yearly_change < 0 Then
            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
        
        End If
          
        'Add one to the summary table row
      summary_table_row = summary_table_row + 1
         
       'Reset the total_stock_volume
      total_stock_volume = 0

    'If the cell on a following a row is the same ticker
    Else
            
       'Add to the total_stock_volume
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                      
    End If

  Next i

ws.Columns("K").NumberFormat = "0.00%"
ws.Columns("I:L").AutoFit

Next ws

End Sub


