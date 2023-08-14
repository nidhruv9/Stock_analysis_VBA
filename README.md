# Stock_analysis_VBA
Submitting the multiple year stock analysis
Sub small_stock_analysis()

Dim total As Double
Dim RIndex As Long
Dim change As Double
Dim CIndex As Integer
Dim start As Long
Dim RCount As Long
Dim percentChange As Double
Dim days As Integer
Dim dailyChange As Single
Dim averageChange As Double
Dim ws As Worksheet

For Each ws In Worksheets
    CIndex = 0
    total = 0
    change = 0
    start = 2
    dailyChange = 0
    
    'set title row
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'get the row number of the last row with data
    
    RCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    For RIndex = 2 To RCount
    
    'printing the results if chnage in ticker name
    
    If ws.Cells(RIndex + 1, 1).Value <> ws.Cells(RIndex, 1).Value Then
    
    'to store the results in a new variable
    
     total = total + ws.Cells(RIndex, 7).Value
     
        If total = 0 Then
        
                'printing  the results
                
            ws.Range("I" & 2 + CIndex).Value = Cells(RIndex, 1).Value
            ws.Range("J" & 2 + CIndex).Value = 0
            ws.Range("K" & 2 + CIndex).Value = "%" & 0
            ws.Range("L" & 2 + CIndex).Value = 0
            
            'to again reach at the starting of the new ticker cell
            
    Else
            If ws.Cells(start, 3) = 0 Then
                For find_value = start To RIndex
                    If ws.Cells(find_value, 3) <> 0 Then
                        start = find_value
                            Exit For
                    End If
                Next find_value
            End If
            
            change = (ws.Cells(RIndex, 6) - ws.Cells(start, 3))
            percentChange = change / ws.Cells(start, 3)
            
            start = RIndex + 1
            
            ws.Range("I" & 2 + CIndex).Value = ws.Cells(RIndex, 1).Value
            ws.Range("J" & 2 + CIndex).Value = change
            ws.Range("J" & 2 + CIndex).NumberFormat = "0.00"
            ws.Range("K" & 2 + CIndex).Value = percentChange
            ws.Range("K" & 2 + CIndex).NumberFormat = "0.00%"
            ws.Range("L" & 2 + CIndex).Value = total
            
            Select Case change
            
              Case Is > 0
                ws.Range("J" & 2 + CIndex).Interior.ColorIndex = 4
              Case Is < 0
                ws.Range("J" & 2 + CIndex).Interior.ColorIndex = 3
              Case Else
                ws.Range("J" & 2 + CIndex).Interior.ColorIndex = 0
                
             End Select
           
        End If
        
        total = 0
        change = 0
        CIndex = CIndex + 1
        days = 0
        dailyChange = 0
        
        Else
            'if Ticker is still the same add results
        
            total = total + ws.Cells(RIndex, 7).Value
        
     
     End If
     
     Next RIndex
     
     'with the use if worksheet function (max and min) calculating the percentage changes in the values
     
     
     ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RCount)) * 100
     ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RCount)) * 100
     ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & RCount))
     
     'storing the value in  variables for greatest and least value
     
     
     increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RCount)), ws.Range("K2:K" & RCount), 0)
     decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RCount)), ws.Range("K2:K" & RCount), 0)
     volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RCount)), ws.Range("L2:L" & RCount), 0)
     
     'applying  the value in the cells "P2","P3" and "P4"
     
     ws.Range("P2") = ws.Cells(increase_number + 1, 9)
     ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
     ws.Range("P4") = ws.Cells(volume_number + 1, 9)

     
     
     
     
     
    
Next ws

End Sub
