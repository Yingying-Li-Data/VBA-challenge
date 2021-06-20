Attribute VB_Name = "Module2"
Sub Summary()

    For Each ws In Worksheets
    
        ws.Cells(2, 15) = "Greatest % increase"
        ws.Cells(3, 15) = "Greatest % decrease"
        ws.Cells(4, 15) = "Greatest total volume"
        
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        
        max_increase = 0
        max_increase_Volume = 0
        max_Decrease = 0
        max_Decrease_Volume = 0
        max_total_volume = 0
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow
        
            If ws.Cells(i, 11).Value > max_increase Then
                max_increase = ws.Cells(i, 11).Value
                max_increase_ticker = ws.Cells(i, 9).Value
                max_increase_Volume = ws.Cells(i, 12).Value
                
            ElseIf Cells(i, 11).Value < max_Decrease Then
                max_Decrease = ws.Cells(i, 11).Value
                max_Decrease_ticker = ws.Cells(i, 9).Value
                max_Decrease_Volume = ws.Cells(i, 12).Value
            
            End If
            
            If Cells(i, 12).Value > max_total_volume Then
                max_total_volume_ticker = ws.Cells(i, 9).Value
                max_total_volume = ws.Cells(i, 12).Value
                
            End If
                            
        Next i
        
        ws.Cells(2, 16).Value = max_increase_ticker
        ws.Cells(2, 17).Value = FormatPercent(max_increase, 2)
        ws.Cells(3, 16).Value = max_Decrease_ticker
        ws.Cells(3, 17).Value = FormatPercent(max_Decrease, 2)
        ws.Cells(4, 16).Value = max_total_volume_ticker
        ws.Cells(4, 17).Value = max_increase_Volume

    Next ws

End Sub
