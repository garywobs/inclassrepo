Attribute VB_Name = "Module1"
Sub StockSymbolVolume()
    
    
'define parameters
Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
vol = 0
Dim Summary_Table_Row As Integer
      
For Each ws In ThisWorkbook.Worksheets
        'set column headers
    ws.Cells(1, 9).Value = "Ticker"
    
    ws.Cells(1, 10).Value = "Total Stock Volume"
    
        'setup integers for loop
    Summary_Table_Row = 2
        
        'loop
    For i = 2 To ws.UsedRange.Rows.Count
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'find all the values
            ticker = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value
                
        'insert values into summary
            ws.Cells(Summary_Table_Row, 9).Value = ticker
                   
            ws.Cells(Summary_Table_Row, 10).Value = vol
                Summary_Table_Row = Summary_Table_Row + 1
                 
                vol = 0
                
        'If the cell immediately following a row is the same brand...
    Else

        'Add to the Brand Total
      vol = vol + ws.Cells(i, 7).Value
        
        End If
        Next i
        'finish loop
    
Next
End Sub
