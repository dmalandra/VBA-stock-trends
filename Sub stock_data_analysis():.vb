Sub stock_data_analysis():

Dim ws As Worksheet

For Each ws In Worksheets

Dim col_idx As Long
Dim i As Long
Dim LastRow As Long
Dim total_volume As Double
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Variant
Dim max_percent As Double
Dim min_percent As Double
Dim max_volume As Double

Dim max_ticker As String
Dim min_ticker As String
Dim vol_ticker As String

col_idx = 2
total_volume = 0

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
open_price = ws.Cells(2, 3).Value

'iterate through all rows

For i = 2 To LastRow
    
    total_volume = total_volume + ws.Cells(i, 7).Value
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(col_idx, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(col_idx, 12).Value = total_volume
        
        close_price = ws.Cells(i, 6).Value
    
        yearly_change = close_price - open_price
        
        percent_change = FormatPercent(yearly_change / open_price)
    
        ws.Cells(col_idx, 10).Value = yearly_change
        ws.Cells(col_idx, 11).Value = percent_change
        
    ElseIf ws.Cells(i, 10) < 0 Then
        ws.Cells(col_idx, 10).Interior.ColorIndex = 3
        ws.Cells(col_idx, 11).Interior.ColorIndex = 3
    
    Else: ws.Cells(col_idx, 10).Interior.ColorIndex = 4
        ws.Cells(col_idx, 11).Interior.ColorIndex = 4
    
    End If

    
    'reset to zero for next ticker
            
        total_volume = 0
        open_price = ws.Cells(i + 1, 3).Value
        close_price = ws.Cells(i + 1, 6).Value
        percent_change = 0
        col_idx = col_idx + 1
         

Next i

max_percent = WorksheetFunction.Max(ws.Columns("K").Value)
min_percent = WorksheetFunction.Min(ws.Columns("K").Value)
max_volume = WorksheetFunction.Max(ws.Columns("L").Value)

ws.Cells(2, 16).Value = max_percent
ws.Cells(3, 16).Value = min_percent
ws.Cells(4, 16).Value = max_volume

For i = 2 To LastRow

    If max_percent = ws.Cells(i, 11).Value Then
        ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
    ElseIf min_percent = ws.Cells(i, 11).Value Then
        ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
    ElseIf max_volume = ws.Cells(i, 12).Value Then
        ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
    End If
    
Next i

    

Next ws

        
End Sub
