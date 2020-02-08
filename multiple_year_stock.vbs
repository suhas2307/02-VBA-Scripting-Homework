Sub multiple_year_stock():
    Dim r, tab_count As Integer
    Dim open_price, close_price, diff_price, tot_vol, high_per_until_now, low_per_until_now, high_tot_vol, last_row As Double
    Dim high_per_tkr, low_per_tkr, high_vol_tkr As String
    Dim ws As Worksheet
    
    For Each ws In Worksheets
        tab_count = 2 ' Reset the table counter for every worksheet
        tot_vol = 0 ' Reset the total volume for each sheet
        high_per_until_now = -1
        low_per_until_now = 1
        high_tot_vol = 0
        
        ' Write the headers for the summary table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ' Determine the last row for each worksheet
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Assign open price for the inital ticker
        open_price = ws.Cells(2, 3).Value
        
        For r = 2 To last_row
            If ws.Cells(r, 1).Value = ws.Cells(r + 1, 1).Value Then
                tot_vol = tot_vol + ws.Cells(r, 7).Value
            Else
                ' Assign close price from the last row for the ticker
                close_price = ws.Cells(r, 6).Value
                ' Calculate the difference in price
                diff_price = close_price - open_price
                ' Assign final values to the table
                ws.Range("I" & tab_count).Value = ws.Cells(r, 1).Value
                ws.Range("J" & tab_count).Value = diff_price
                ' if open price is not 0 then divide the difference in price by open price to get the percentage change. Else display NaN
                If open_price <> 0 Then
                    percent_change = diff_price / open_price
                    ws.Range("K" & tab_count).Value = percent_change
                Else
                    ws.Range("K" & tab_count).Value = "NaN"
                End If
                
                ws.Range("K" & tab_count).NumberFormat = "0.00%" ' Change format to percentage
                ' Fill color based on the difference in price
                If diff_price >= 0 Then
                    ws.Range("J" & tab_count).Interior.ColorIndex = 4
                    ' To get the greatest percent increase
                    If percent_change > high_per_until_now And open_price <> 0 Then
                        high_per_until_now = percent_change
                        high_per_tkr = ws.Cells(r, 1).Value
                    End If
                Else
                    ws.Range("J" & tab_count).Interior.ColorIndex = 3
                    ' To get the greatest percent decrease
                    If percent_change < low_per_until_now Then
                        low_per_until_now = percent_change
                        low_per_tkr = ws.Cells(r, 1).Value
                    End If
                End If
                
                tot_vol = tot_vol + ws.Cells(r, 7).Value
                ws.Range("L" & tab_count).Value = tot_vol
                
                'Check for greatest total volume
                If tot_vol > high_tot_vol Then
                    high_tot_vol = tot_vol
                    high_vol_tkr = ws.Cells(r, 1).Value
                End If
                
                ' Reset total volume
                tot_vol = 0
                'Assign open_price for the next iteration
                open_price = ws.Cells(r + 1, 3).Value
                'Increment table counter for the next stock ticker
                tab_count = tab_count + 1
            End If
        Next r
        
        ' Write the greatest stat table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P2").Value = high_per_tkr
        ws.Range("P3").Value = low_per_tkr
        ws.Range("P4").Value = high_vol_tkr
        If high_per_until_now <> -1 Then
            ws.Range("Q2").Value = high_per_until_now
        End If
        If low_per_until_now <> 1 Then
            ws.Range("Q3").Value = low_per_until_now
        End If
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = high_tot_vol
        ws.Range("Q4").NumberFormat = "0.0000E+00"
        
        ws.Columns("I:Q").AutoFit
    Next ws
End Sub


