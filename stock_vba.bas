Attribute VB_Name = "Module1"
Sub stocks()

    'create variables for given data
    Dim i As Double
    Dim ticker As String
    Dim open_price As Double
    Dim close_price As Double

    'create variables for calculated data
    Dim sum_row As Integer
    Dim yearly_change As Double
    Dim total_stock_volume As Double
    Dim percent_change As Double
    
    For Each ws In Worksheets
    
        'add headers for new columns
        ws.Range("J1").Value = "Ticker Symbol"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"
    
        'identify last row in worksheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'set row for summary table to 1
        sum_row = 1
        
        'loop through the rows of data
        For i = 2 To last_row
            
            total_stock_volume = 0
            
            'if ticker is unique (first row of the stock)
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                
                'variables
                open_price = ws.Cells(i, 3).Value
                total_stock_volume = ws.Cells(i, 7).Value
                ticker = ws.Cells(i, 1).Value
                
                'calculations
                sum_row = sum_row + 1
                
                'add ticker to summary table
                ws.Cells(sum_row, 10).Value = ticker
                
            'if the next ticker is unique (last row of stock)
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                'variables
                close_price = ws.Cells(i, 6).Value
                
                'calculations
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                yearly_change = close_price - open_price
                percent_change = (ws.Cells(i, 6).Value - open_price) / open_price
                
                'add calculations to summary table
                ws.Cells(sum_row, 11).Value = yearly_change
                ws.Cells(sum_row, 13).Value = total_stock_volume
                ws.Cells(sum_row, 12).Value = percent_change
                
                'color and format summary table
                If yearly_change < 0 Then
                    ws.Cells(sum_row, 11).Interior.ColorIndex = 3
                ElseIf yearly_change > 0 Then
                    ws.Cells(sum_row, 11).Interior.ColorIndex = 4
                End If
                ws.Cells(sum_row, 12).NumberFormat = "0.00%"
          
            End If
            
        Next i
        
        
            
    
    Next ws
    
        
End Sub
