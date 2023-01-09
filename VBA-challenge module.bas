Attribute VB_Name = "Module1"

Sub StockAnalysis()

 For Each ws In Worksheets
 
 'set variables
 Dim ticker As String
 
 Dim tickervolume As Double
 tickervolume = 0
 
 Dim tickerrow As Integer
 tickerrow = 2
 
 
 'open price
 Dim oprice As Double
 
 oprice = ws.Cells(2, 3).Value
 
 'close price
 Dim cprice As Double
 'yearly change
 Dim ychange As Double
 'percent change
 Dim pchange As Double
 
 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = " Total Stock Volume"
 
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 For i = 2 To lastrow
 
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker = ws.Cells(i, 1).Value
    
    tickervolume = tickervolume + ws.Cells(i, 7).Value
    
    ws.Range("I" & tickerrow).Value = ticker
    
    ws.Range("L" & tickerrow).Value = tickervolume
    
    'collect closing price
    cprice = ws.Cells(i, 6).Value
    
    'yearly change calculations
    ychange = (cprice - oprice)
    
    ws.Range("J" & tickerrow).Value = ychange
    
        If oprice = 0 Then
            pchange = 0
        Else
            pchange = ychange / oprice
        End If
        
    ws.Range("K" & tickerrow).Value = pchange
    ws.Range("K" & tickerrow).NumberFormat = "0.00%"
    
    tickerrow = tickerrow + 1
    
    tickervolume = 0
    
    oprice = ws.Cells(i + 1, 3)
Else

    tickervolume = tickervolume + ws.Cells(i, 7).Value
    
    
End If

Next i
            ' Conditional formatting
        lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For i = 2 To lastrow_summary_table
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 10
            
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If

 Next i

    
    'label the cells

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    ' "Percent Change" max and min value and max in "Total Stock Volume"
    
        For i = 2 To lastrow_summary_table
        
            'maximum percent change
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"

            'minimum percent change
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow_summary_table)) Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            
            'maximum volume of trade
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow_summary_table)) Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
    
    Next ws
        
        

 
 
 
 
End Sub
