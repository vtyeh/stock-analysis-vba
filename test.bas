Sub LoopStocks()
'Loop through all worksheets
For Each ws In Worksheets

    'Set variable for for loop
    Dim j As Long
    'Set variable to keep track of ticker
    Dim ticker As String
    
     'Define loop for hard part
    Dim i As Double
    'Set variable to keep track of ticker
    Dim hardTicker As String
    
    'Set variable to keep track of total volume
    Dim sumVol As Double
        sumVol = 0
    
    'Set variable to keep track of ticker row for stock summary
    Dim tickSym As Double
        tickSym = 2
    
    'Set variables for yearly change
    Dim open_price As Double
    Dim close_price As Double
    Dim change As Double
    
    'Find last row for j
    Dim LastRow As Double
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    'to speed things up, it's good to minimize lookups from the table
    'create a couple more strings to save (j,1) and (j +1,1) since they're referenced multiple times
    Dim one, two As String
    
    'Begin loop
    For j = 2 To LastRow
        
    'save (j,1) and (j+1,1) at the beginning of the loop
    current_one = ws.Cells(j, 1).Value
    next_one = ws.Cells(j + 1, 1).Value
    previous_one = ws.Cells(j - 1, 1).Value
    
    'Find total volume (this has to be run everytime, so it can be outside any if/else statement)
    sumVol = sumVol + ws.Cells(j, 7).Value
            
            'In this case we can combine our logic into a single if/else to reduce conditionals
            If current_one <> next_one Then
                close_price = ws.Cells(j, 6).Value
                
                'If the loop gets to this point, the open price has been retrieved
                change = close_price - open_price
                
                'Create a nested If statement to calculate the percent change with the open price and close price
                ws.Cells(tickSym, 10).Value = change
                    If open_price <> 0 Then 'The divisor cannot be 0 but the numerator can
                        percentage = change / open_price
                        ws.Range("K" & tickSym).Value = percentage
                    End If
                
                'Place total volume and ticker to stock summary
                ws.Range("I" & tickSym).Value = current_one
                ws.Range("L" & tickSym).Value = sumVol
                'Add a row for next ticker
                tickSym = tickSym + 1
                'Reset volume to 0 everytime a new ticker is found'
                sumVol = 0
                
             ElseIf previous_one <> current_one Then
                open_price = ws.Cells(j, 3).Value
             End If
        Next j
        
        'Move formatting outside the of the loop
        ws.Range("K:K").EntireColumn.NumberFormat = "0.00%"
        ws.Range("L:L").EntireColumn.NumberFormat = "0"
    
   
    
    'Find the greatest increase, decrease, and volume
    greatest_increase = Application.WorksheetFunction.Max(ws.Range("K:K"))
    ws.Cells(2, 16).Value = greatest_increase
    greatest_decrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
    ws.Cells(3, 16).Value = greatest_decrease
    greatest_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(4, 16).Value = greatest_volume
    
    'Begin loop
        For i = 2 To LastRow
        condForm = ws.Cells(i, 11).Value
        hardTicker = ws.Cells(i, 9).Value
        
            If ws.Cells(i, 11).Value = greatest_increase Then
                ws.Cells(2, 15).Value = hardTicker
                
            ElseIf ws.Cells(i, 11).Value = greatest_decrease Then
                ws.Cells(3, 15).Value = hardTicker
                
            ElseIf ws.Cells(i, 12).Value = greatest_volume Then
                ws.Cells(4, 15).Value = hardTicker
                
            End If
            
            'Conditionally format the percent change
            If condForm < 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 38
            ElseIf condForm > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 35
            End If
    Next i
    'Change P2 and P3 to percent
    ws.Range("P2:P3").NumberFormat = "0.00%"

    'Format worksheets
        'Make column titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'Make row titles
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        
        'Format titles
        ws.Range("I1").Font.Bold = True
        ws.Range("J1").Font.Bold = True
        ws.Range("K1").Font.Bold = True
        ws.Range("L1").Font.Bold = True
        ws.Range("O1").Font.Bold = True
        ws.Range("P1").Font.Bold = True
        ws.Range("N2").Font.Bold = True
        ws.Range("N3").Font.Bold = True
        ws.Range("N4").Font.Bold = True
        
        ws.Range("A:P").EntireColumn.AutoFit
Next ws
End Sub