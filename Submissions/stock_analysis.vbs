Option Explicit

Sub stock_analysis()
    Dim row As Double
    Dim ticker As String ' Values for new column
    Dim quarterly_change As Double ' Values for new column
    Dim percent_change As Double ' Values for new column
    Dim total_stock_volume As Double ' Values for new column
    Dim final_row As Double
    Dim nt As Double ' row counter for new table entries
    Dim stock As Double ' current stock volume
    Dim open_value As Double ' open stock value
    Dim close_value As Double ' close stock value
    Dim greatest_increase As Double ' Greatest % Increase value
    Dim greatest_decrease As Double ' Greatest % Decrease value
    Dim greatest_volume As Double ' Greatest Total Volume value
    Dim greatest_increase_ticker As String
    Dim greatest_decrease_ticker As String
    Dim greatest_volume_ticker As String
    Dim ws As Worksheet
    
    ' Disable Real-time updating for faster code runtime
    Application.ScreenUpdating = False
    
    For Each ws In Worksheets
        final_row = ws.Cells(1, 1).End(xlDown).row
        total_stock_volume = 0
        nt = 2
        open_value = ws.Cells(2, 3).Value
    
        ' Prepare new table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        For row = 2 To final_row
        
            ticker = ws.Cells(row, 1).Value
            stock = ws.Cells(row, 7).Value
            
            If ws.Cells(row + 1, 1).Value <> ticker Then
                ' new ticker value detected, output current values
                
                close_value = ws.Cells(row, 6).Value
                
                quarterly_change = close_value - open_value
                percent_change = ((close_value - open_value) / open_value)
                
                total_stock_volume = total_stock_volume + stock
                ws.Cells(nt, 9).Value = ticker
                ws.Cells(nt, 12).Value = total_stock_volume
                ws.Cells(nt, 10).Value = quarterly_change
                ws.Cells(nt, 11).Value = FormatPercent(percent_change, 2)
                ws.Cells(nt, 10).NumberFormat = "#,##0.00"
                
                If quarterly_change > 0 Then
                
                    ws.Cells(nt, 10).Interior.Color = vbGreen
                    ws.Cells(nt, 11).Interior.Color = vbGreen
                    
                ElseIf quarterly_change < 0 Then
                    
                    ws.Cells(nt, 10).Interior.Color = vbRed
                    ws.Cells(nt, 11).Interior.Color = vbRed
                
                End If
                
                ' reset
                nt = nt + 1
                total_stock_volume = 0
                open_value = ws.Cells(row + 1, 3).Value
            
            Else
            
                total_stock_volume = total_stock_volume + stock
            
            End If
        
        Next row
        
        ' Prepare new table values
        final_row = ws.Cells(1, 11).End(xlDown).row
        greatest_increase = ws.Cells(2, 11).Value
        greatest_decrease = ws.Cells(2, 11).Value
        greatest_volume = ws.Cells(2, 12).Value
        
        For row = 2 To final_row
        
            If ws.Cells(row, 11).Value > greatest_increase Then
            
                greatest_increase = ws.Cells(row, 11).Value
                greatest_increase_ticker = ws.Cells(row, 9).Value
            
            End If
            
            If ws.Cells(row, 11).Value < greatest_decrease Then
            
                greatest_decrease = ws.Cells(row, 11).Value
                greatest_decrease_ticker = ws.Cells(row, 9).Value
                
            End If
            
            If ws.Cells(row, 12).Value > greatest_volume Then
            
                greatest_volume = ws.Cells(row, 12).Value
                greatest_volume_ticker = ws.Cells(row, 9).Value
                
            End If
        
        Next row
        
        ' Output values for new table
        ws.Cells(2, 17).Value = FormatPercent(greatest_increase, 2)
        ws.Cells(3, 17).Value = FormatPercent(greatest_decrease, 2)
        ws.Cells(4, 17).Value = greatest_volume
        ws.Cells(2, 16).Value = greatest_increase_ticker
        ws.Cells(3, 16).Value = greatest_decrease_ticker
        ws.Cells(4, 16).Value = greatest_volume_ticker
        
        ws.Columns("J:L").AutoFit
        ws.Columns("O:Q").AutoFit
    Next ws

    ' Enable screen updating now that code is complete
    Application.ScreenUpdating = True
    MsgBox ("Calculations Complete")

End Sub


Sub Reset()

    Dim ws As Worksheet
    
    For Each ws In Worksheets

        ws.Cells(1, 9).CurrentRegion.Clear
        ws.Cells(1, 15).CurrentRegion.Clear
    
    Next ws

End Sub

