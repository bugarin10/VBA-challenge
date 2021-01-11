' This macro generates two summary tables
' The columns of the first table are: Ticker, YearlyChange, Percentage Change, Total Stock Volume
' The second one will have the greatest and smallest changes in percetage and the greatest total volume

Sub VBA()


For Each ws In Worksheets
'''''''' First table ''''''''''''''''''''''''''''''''

    ws.Columns("K").NumberFormat = "0.00%"

    Dim n_tickers As Long
    Dim nrows As Long
    Dim total_stock As LongLong
    Dim abs_change As Double
    ' Defining first opening and last close
    Dim ope As Single
    Dim clo As Single
    Dim ticker As String
    
    nrows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Assigning headers
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percentage Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' Counter for rows in table 1
    n_tickers = 1
    ' Current ticker
    ticker = ws.Cells(2, 1).Value
    ' the open stars in the second row
    ope = ws.Cells(2, 3).Value
    ' Setting the first ticker
    ws.Cells(2, 9).Value = ticker
    ' Adding stock
    total_stock = ws.Cells(2, 7).Value
    
    
    For i = 3 To nrows
    
    total_stock = total_stock + ws.Cells(i, 7).Value
    
    ''''''''''''''''' If we find different sticker then '''''''''''''''''''''
    
        If ws.Cells(i, 1).Value <> ticker Then
            
            'closing is the last one
            clo = ws.Cells(i - 1, 6).Value
            'we have new ticker
            ticker = ws.Cells(i, 1).Value
            'Yearly abs change
            abs_change = clo - ope
            ws.Cells(n_tickers + 1, 10).Value = abs_change
            'Pct Change
            If ope > 0 Then
                ws.Cells(n_tickers + 1, 11).Value = abs_change / ope
            Else
                ws.Cells(n_tickers + 1, 11).Value = 0
            End If
            
            'New open
            ope = ws.Cells(i, 3).Value
            'Total Stock
            ws.Cells(n_tickers + 1, 12).Value = total_stock - ws.Cells(i, 7).Value
            'Re-start total stock
            total_stock = ws.Cells(i, 7).Value
            
            ' Changing colors
                If abs_change >= 0 Then
                    ws.Cells(n_tickers + 1, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(n_tickers + 1, 10).Interior.ColorIndex = 3
                End If
                
                
            'change to next row in the new table
            n_tickers = n_tickers + 1
            ' We have new ticker
            ws.Cells(n_tickers + 1, 9).Value = ticker
            ' We have new open
            ope = ws.Cells(i, 3).Value
        
        End If
        
        If i = nrows Then
            clo = ws.Cells(i, 6).Value
            'Yearly abs change
            abs_change = clo - ope
            ws.Cells(n_tickers + 1, 10).Value = abs_change
            'Pct Change
            If ope > 0 Then
                ws.Cells(n_tickers + 1, 11).Value = abs_change / ope
            Else
                ws.Cells(n_tickers + 1, 11).Value = 0
            End If
            'Total Stock
            ws.Cells(n_tickers + 1, 12).Value = total_stock
            
            ' Changing colors
                If abs_change >= 0 Then
                    ws.Cells(n_tickers + 1, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(n_tickers + 1, 10).Interior.ColorIndex = 3
                End If
            
        End If
           
    Next i
    
    
    ''''''''''''''''''' Second Table ''''''''''''''''''''''''
    
    Dim min_change As Double
    Dim max_change As Double
    Dim ticker_min As String
    Dim ticker_max As String
    Dim ticker_stock As String
    Dim max_stock As LongLong
    
    
    'Setting values
    min_change = ws.Cells(2, 11).Value
    max_change = ws.Cells(2, 11).Value
    max_stock = ws.Cells(2, 12).Value
    ticker_min = ws.Cells(2, 9).Value
    ticker_max = ws.Cells(2, 9).Value
    ticker_stock = ws.Cells(2, 9).Value
    
    
    For i = 2 To n_tickers - 1:
        
        If min_change >= ws.Cells(i + 1, 11).Value Then
            min_change = ws.Cells(i + 1, 11).Value
            ticker_min = ws.Cells(i + 1, 9).Value
       End If
       
       If max_change <= ws.Cells(i + 1, 11).Value Then
            max_change = ws.Cells(i + 1, 11).Value
            ticker_max = ws.Cells(i + 1, 9).Value
       End If
        
       If max_stock <= ws.Cells(i + 1, 12).Value Then
            max_stock = ws.Cells(i + 1, 12).Value
            ticker_stock = ws.Cells(i + 1, 9).Value
       End If
       
    Next i
    
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(2, 17).Value = ticker_max
    ws.Cells(2, 18).Value = max_change
    ws.Cells(2, 18).NumberFormat = "0.00%"
    
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(3, 17).Value = ticker_min
    ws.Cells(3, 18).Value = min_change
    ws.Cells(3, 18).NumberFormat = "0.00%"
    
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    ws.Cells(4, 17).Value = ticker_stock
    ws.Cells(4, 18).Value = max_stock
    
        
 Next ws
    
End Sub
