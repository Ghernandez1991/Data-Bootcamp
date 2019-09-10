Sub ()

Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

'Keep track of the location for each ticker name in the summary column
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Look through all stock tickers
For i = 2 To 760192

' check if we are still within the same ticker name, if it is not...
If Cells(i + 1, 1).Value <> Cells(i, 1) Then

'set the Ticker name
Ticker = Cells(i, 1).Value

'Add to the volume
Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

'Print the Total Stock Volume in the Total Stock Volume Column
Range("I" & Summary_Table_Row).Value = Ticker

'Print the Total Stock Volume in the Total Stock Volume Column
Range("J" & Summary_Table_Row).Value = Total_Stock_Volume

'Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

'Reset the Ticker Total
Total_Stock_Volume = 0

'if the cell immediately following a row is the same brand...
Else
    'Add to the Total_Stock_Volume
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
    End If
Next i

   

End Sub
