Attribute VB_Name = "Module1"
Sub WallStreet()

For Each ws In Worksheets

    'Create the column headers
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Yearly Change"
    ws.Cells(1, 11) = "Percent Change"
    ws.Cells(1, 12) = "Total Stock Volume"

    'Define variables
    Dim Ticker As String
    
    earlestopen = 0
    LatestClose = 0
    volume = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker_row = 2
    
    'Grab unique tickers
    'Calculate total stock volume total = total + vol if ...
    'Calculate yearly change (latest close - earlest open)
    'Calculate percent change ( yearly change / earlest open)
    
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            Ticker = ws.Cells(i, 1).Value
            volume = volume + ws.Cells(i, 7).Value
            
            ws.Cells(ticker_row, 9).Value = Ticker
            ws.Cells(ticker_row, 12).Value = volume
            
            LatestClose = ws.Cells(i, 6).Value
                       
            Change = LatestClose - earlestopen
            
            If earlestopen = 0 Then
                change_percent = ""
                Else: change_percent = FormatPercent(Change / earlestopen, 2)
                
            End If
                      
            ws.Cells(ticker_row, 10).Value = Change
            ws.Cells(ticker_row, 11).Value = change_percent
            
            If ws.Cells(ticker_row, 10) > 0 Then
        
                ws.Cells(ticker_row, 10).Interior.ColorIndex = 4
            
            ElseIf ws.Cells(ticker_row, 10) < 0 Then
            
                ws.Cells(ticker_row, 10).Interior.ColorIndex = 3
                
            Else: ws.Cells(ticker_row, 10).Interior.ColorIndex = 0
        
            End If
            
            ticker_row = ticker_row + 1
            volume = 0
            
        Else
                      
            volume = volume + ws.Cells(i, 7).Value
            
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then earlestopen = ws.Cells(i, 3).Value
            
        End If
        
    Next i
    
    ws.Columns("I:L").AutoFit

Next ws
    
End Sub
