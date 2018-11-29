Sub stockdata():

'loop through every worksheet
Dim ws As Worksheet
For Each ws In Worksheets
    
    'activating worksheets
    ws.Activate
    
    'calculate the last row on each sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'sets ticker header
    ws.Cells(1, 9).Value = "Ticker"
    'sets total volume header
    ws.Cells(1, 10).Value = "Total Stock Volume"
    
    
    'designate ticker variable
    Dim stock_name As String
 
    'designate stock volume variable
    Dim stock_volume As Double
    stock_volume = 0
    
    'designate location
    Dim summary_table As Double
    summary_table = 2
    
    'create for loop to go through the ticker values
    For i = 2 To LastRow
    
        'searches for when the value changes
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'assigns value to stock_name
            stock_name = Cells(i, 1).Value
            'prints ticker
            Range("I" & summary_table).Value = stock_name
            'adds to stock volume
            stock_volume = stock_volume + Cells(i, 7).Value
            'prints ticker and stock volume
            Range("J" & summary_table).Value = stock_volume
            'add a row to the summary
            summary_table = summary_table + 1
            'reset stock volume
            stock_volume = 0
        Else
            'adds total volume
            stock_volume = stock_volume + Cells(i, 7).Value
        End If
        
    Next i
    
Next ws

End Sub


