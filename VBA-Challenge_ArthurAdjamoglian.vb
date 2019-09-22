Sub stock()

    Dim ws As Worksheet
    
    For Each ws In Worksheets

        'define vars
        Dim LR As Long
        Dim ticker As String
        Dim open_price As Single
        Dim close_price As Single
        Dim summary_table_row As Integer
        Dim yearly_change As Single
        Dim stock_vol As Double
        Dim percent_change As Single
        Dim result As Double
        Dim ticker_count As Integer
        Dim start_row As Integer
        
        
        'set var starting values
        summary_table_row = 2
        stock_vol = 0
        start_row = 2
        
        'define last row
        LR = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'create headers for sum_table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        'set first open price
        open_price = ws.Cells(2, 3).Value
            
        'loop through rows 1 to last
        For i = 2 To LR
        
        
            'if the ticker symbol has changed
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            
                'calculate volume of stock that adds all stock volumes for a specific ticker
                stock_vol = stock_vol + ws.Cells(i, 7).Value
                
                'i last row before change in ticker symbol contains close price for that company
                close_price = ws.Cells(i, 6).Value
                 
                'calculate yearly change
                yearly_change = close_price - open_price
                
                'if open_price = 0
                If open_price = 0 Then
                
                If ws.Cells(start_row, 3).Value = 0 Then
                    For j = start_row To i
                        If ws.Cells(start_row, 3).Value <> 0 Then
                            start_row = j
                            percent_change = ((close_price - open_price) / open_price)
                        End If
                        
                    Next j
                    
                
                 End If
                 
                 Else
                 percent_change = ((close_price - open_price) / open_price)
                 End If
                
                
                ws.Range("J" & summary_table_row).Value = yearly_change
                
                    If ws.Range("J" & summary_table_row).Value >= 0 Then
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                    Else
                        ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                    End If
                        
                
                ws.Range("K" & summary_table_row).Value = percent_change
                
                'i+1 (new)ticker symbol row contains open price for that company
                open_price = ws.Cells(i + 1, 3).Value
            
                'set ticker
                ticker = ws.Cells(i, 1).Value
                
                'print tickers into summary table
                ws.Range("I" & summary_table_row).Value = ticker
                
                'print stock_vol into summary table
                ws.Range("L" & summary_table_row).Value = stock_vol
                
                'changes summary table row after each change in ticker value
                summary_table_row = summary_table_row + 1
                
                
                If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
                ticker_count = ticker_count + 1
                End If
                        
                'reset stock_vol
                stock_vol = 0
                
            Else
                stock_vol = stock_vol + ws.Cells(i, 7).Value
                
            End If
            
        Next
        
        
        'calculate and print greatest % increase
        ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 15).Value = ws.Cells(ws.Range("K:K").Find(WorksheetFunction.Max(ws.Range("K:K"))).Row, 9).Value
        
        'calculate and print greatest % decrease
        ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(3, 15).Value = ws.Cells(ws.Range("K:K").Find(WorksheetFunction.Min(ws.Range("K:K"))).Row, 9).Value
        
        'calculate and print greatest total volume
        ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Range("L:L"))
        ws.Cells(4, 15).Value = ws.Cells(ws.Range("L:L").Find(WorksheetFunction.Max(ws.Range("L:L"))).Row, 9).Value
        
        ws.Range("K2:K" & summary_table_row).NumberFormat = "0.00%"
        ws.Range("J2:J" & summary_table_row).NumberFormat = "0.00000000"
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
    Next ws

End Sub
