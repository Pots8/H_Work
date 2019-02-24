Sub findeasy()

' loop all worksheets
For Each ws In Worksheets

Dim Ticker_Name As String
Dim Total As Double
Ticker_Total = 0

' Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
'Keep track of ticker in summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'add title row
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

'do a loop for all rows

For i = 2 To LastRow
'For i = 2 To 70926


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       
    Ticker_Name = ws.Cells(i, 1).Value
    
     'add to total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
    'print ticker name
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    
    'print total
    ws.Range("J" & Summary_Table_Row).Value = Ticker_Total
    
   
     
    'move to next summary row
    Summary_Table_Row = Summary_Table_Row + 1
    
    
    
    'make total to zero
    Ticker_Total = 0
    
    ' if still same ticker continue with adding
    Else
    
    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
    
        
    End If
    
 Next i
    
Next ws

End Sub
