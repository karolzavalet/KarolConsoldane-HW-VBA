Sub Ticker_Names()
Dim last_row As Long, c As Integer, g_dec As Double, g_inc As Double
    For Each ws In Worksheets
        'labels
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volumen"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        'Value of the row the new table is
        c = 2
        'total_vol initial value is zero
        total_vol = 0
        'calculate the last row
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Iterate in the ticker column to get the different ticker names and sum of vol
        For i = 2 To last_row
            'If the current row (i) name is different to the following row (i+1)...
            'then we have a new ticker name
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'Get the ticker name
                ws.Cells(c, 9).Value = ws.Cells(i, 1).Value
                'This will add up the last value of "<vol>" to the total_vol sum
                total_vol = total_vol + ws.Cells(i, 7).Value
                'Print the amount of all the sum per ticker name
                ws.Cells(c, 12).Value = total_vol
                'if the name is different it should move on the nex row
                c = c + 1
                'for the new ticker we have initialize the sum of "<vol>"
                total_vol = 0
            Else
                'this is adding all the "<vol>" of the same tickers name when...
                'the current row (i) name is different to the following row (i+1)
                total_vol = total_vol + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub


Sub SecondChart()
Dim FMin As Long, FMax As Long, last_row As Long, open_val As Double 
Dim close_val As Double, year_chg As Double, perc_chg As Double
    For Each ws In Worksheets
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'k would be value of d' row where the new ticker starts
        k = 2
        'c is the row where each ticker is in 2 chart
        c = 2
        For i = 2 To last_row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Rng = ws.Range("B" & k, "B" & i)
                'Get the earliest date
                FMin = Application.WorksheetFunction.Min(Rng)
                'Get the last date
                FMax = Application.WorksheetFunction.Max(Rng)
                For m = k To i
                    If FMin = ws.Cells(m, 2).Value Then
                        open_val = ws.Cells(m, 3).Value
                    End If
                    
                    If FMax = ws.Cells(m, 2).Value Then
                        close_val = ws.Cells(m, 6).Value
                    End If
                Next m

                'calculate yearly change
                year_chg = close_val - open_val
                'Print "yearly change"
                ws.Cells(c, 10).Value = year_chg
                If year_chg < 0 Then
                    ws.Range("J" & c).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & c).Interior.ColorIndex = 4
                End If
                    
                'Apply percentage format to the "percentage change" column
                ws.Cells(c, 11).NumberFormat = "0.00%"
                'Print and calculate "percentage change"
                If open_val = 0 Then
                    perc_chg = close_val
                    ws.Cells(c, 11) = perc_chg
                Else
                    perc_chg = (year_chg / open_val)
                    ws.Cells(c, 11) = (year_chg / open_val)
                End If
            year_chg = 0
            perc_chg = 0
            c = c + 1
            k = i + 1
            End If
        Next i
    Next ws
End Sub


Sub third_chart()
Dim last_row As Long, c As Integer, g_dec As Double, g_inc As Double, last_row2 As Long
    For Each ws In Worksheets
        'for the greatest increases, decreases, etc
        'number of rows in the second chart
        last_row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        Rng2 = ws.Range("K2", "K" & last_row2)
        Rng3 = ws.Range("L2", "L" & last_row2)
        g_inc = Application.WorksheetFunction.Max(Rng2)
        g_dec = Application.WorksheetFunction.Min(Rng2)
        g_tot = Application.WorksheetFunction.Max(Rng3)
        'Apply percentage format to the "Greatest increase and decrease"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        'Print the gratest value
        ws.Cells(2, 17) = g_inc
        ws.Cells(3, 17) = g_dec
        ws.Cells(4, 17) = g_tot
        'Get the tickers names of the greatest values
        For a = 2 To last_row2
            If g_inc = ws.Cells(a, 11).Value Then
                ws.Cells(2, 16).Value = ws.Cells(a, 9).Value
            End If
            If g_dec = ws.Cells(a, 11).Value Then
                ws.Cells(3, 16).Value = ws.Cells(a, 9).Value
            End If
            If g_tot = ws.Cells(a, 12).Value Then
                ws.Cells(4, 16).Value = ws.Cells(a, 9).Value
            End If
        Next a
    Next ws
End Sub

