Attribute VB_Name = "Module1"
Sub Multiple_Stock_Analysis()

' Loop through year 2014-2016 worksheets
    
        For Each ws In Worksheets
    
' Column Headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
 ' Declare variables
       
        Dim ticker_symbol As String
        Dim total_vol As Double
        total_vol = 0
        Dim rowcount As Long
        rowcount = 2
        Dim year_open As Double
        year_open = 0
        Dim year_close As Double
        year_close = 0
        Dim year_change As Double
        year_change = 0
        Dim percent_change As Double
        percent_change = 0
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop ticker symbols
        For i = 2 To lastrow
            
            'Conditionals
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

                year_open = ws.Cells(i, 3).Value

            End If

         
            total_vol = total_vol + ws.Cells(i, 7)

            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(rowcount, 12).Value = total_vol
                year_close = ws.Cells(i, 6).Value
                year_change = year_close - year_open
                ws.Cells(rowcount, 10).Value = year_change
                
                If year_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                    
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                    
                End If

                'percent change for summary table
                If year_open = 0 And year_close = 0 Then
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                    
                ElseIf year_open = 0 Then
                    
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    ws.Cells(rowcount, 11).Value = percent_change
                    
                Else
                    percent_change = year_change / year_open
                    ws.Cells(rowcount, 11).Value = percent_change
                    ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                    
                End If

              
                rowcount = rowcount + 1

                'Reset variables
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
                
            End If
            
        Next i

       
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

       
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

      
        Dim best_stock As String
        Dim best_value As Double

        
        best_value = ws.Cells(2, 11).Value

        Dim worst_stock As String
        Dim worst_value As Double

       
        worst_value = ws.Cells(2, 11).Value

        Dim most_vol_stock As String
        Dim most_vol_value As Double

      
        most_vol_value = ws.Cells(2, 12).Value

        'Loop through summary table
        For j = 2 To lastrow

            'Conditional for best performer
            If ws.Cells(j, 11).Value > best_value Then
                best_value = ws.Cells(j, 11).Value
                best_stock = ws.Cells(j, 9).Value
                
            End If

            'Conditional for worst performer
            If ws.Cells(j, 11).Value < worst_value Then
                worst_value = ws.Cells(j, 11).Value
                worst_stock = ws.Cells(j, 9).Value
                
            End If

            'greatest volume
            If ws.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = ws.Cells(j, 12).Value
                most_vol_stock = ws.Cells(j, 9).Value
            End If

        Next j

        
        ws.Cells(2, 16).Value = best_stock
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_stock
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value

        'Autofit
        ws.Columns("I:L").EntireColumn.AutoFit
        ws.Columns("O:Q").EntireColumn.AutoFit

    Next ws

End Sub



