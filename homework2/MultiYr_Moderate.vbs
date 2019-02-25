Sub moderate()

For Each ws In Worksheets

Dim Ticker_Name As String
Dim Total As Double
Dim Start_p As Single
Dim Close_p As Single
Dim Test As Boolean
Ticker_Total = 0
' Determine the Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
'Keep track of ticker in summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
Start_p = 0
Test = False

'add title row
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'do a loop

For i = 2 To LastRow


    If Test = False Then
        Start_p = ws.Cells(i, 3).Value
        Test = True
    End If
    
    'check if ticker change

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    Ticker_Name = ws.Cells(i, 1).Value
    
         
     Test = False
     Close_p = ws.Cells(i, 6).Value
     year_change = Close_p - Start_p
     If Start_p = 0 Then
         year_ratio = 0
     Else
         year_ratio = Round((year_change / Start_p * 100), 2)
     End If
     Start_p = 0
     
     'add to total
      Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
    'print ticker name
    ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
    
    'print total
    ws.Range("J" & Summary_Table_Row).Value = year_change
    
    'print total
    ws.Range("K" & Summary_Table_Row).Value = year_ratio & "%"
    
    'print total
    ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
        
   'change to green color when positive, red color when negative
     
     Select Case year_change
    
    Case Is > 0
    
     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
    Case Is < 0
     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
     
     Case Else
     
     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
    
    
    End Select
     
     
    'move to next summary row
    Summary_Table_Row = Summary_Table_Row + 1
    
        
    'make total to zero
    Ticker_Total = 0
    
        
    Else
    
    Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
    
            
    End If
    
        
 Next i
    
Next ws

End Sub


