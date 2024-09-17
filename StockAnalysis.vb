Sub StockAnalysis()

    ' Variable declarations
    Dim ws As Worksheet
    Dim res As Worksheet
    Dim last_row As Long
    Dim i As Long
    Dim ticker As String
    Dim curr_ticker As String
    Dim prev_ticker As String
    Dim s_price As Double
    Dim e_price As Double
    Dim volume As Double
    Dim volumeForTicker As Double
    Dim per_change As Double
    Dim quart_start As Long
    Dim quart_end As Long
    Dim curr_quart As String
    Dim prev_quart As String
    
    Dim os_price As Double
    Dim oe_price As Double
    Dim o_per_change As Double
 
    Dim great_incr As Double
    Dim great_decr As Double
    Dim great_vol As Double
    Dim great_incr_ticker As String
    Dim great_decr_ticker As String
    Dim great_vol_ticker As String
    
    ' Initialize greatest values
    great_incr = -999999
    great_decr = 999999
    great_vol = 0
    
    ' Set up 'Results' sheet
    On Error Resume Next
    Set res = ThisWorkbook.Sheets("Results")
    If res Is Nothing Then
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = "Results"
        Sheets("Results").Activate
        Set res = ThisWorkbook.Sheets("Results")
    End If
    On Error GoTo 0
    
    ' Clear 'Results' sheet before re-use
    res.Cells.Clear
    
    ' Headers in the 'Results' sheet for table 1
    res.Cells(1, 1).Value = "Ticker"
    res.Cells(1, 2).Value = "Quarter"
    res.Cells(1, 3).Value = "Quarterly Change"
    res.Cells(1, 4).Value = "Percentage Change"
    res.Cells(1, 5).Value = "Total Stock Volume"
    
    ' Headers for table 2
    res.Cells(1, 9).Value = "Ticker"
    res.Cells(1, 10).Value = "Value"
    res.Cells(2, 8).Value = "Greatest % Increase"
    res.Cells(3, 8).Value = "Greatest % Decrease"
    res.Cells(4, 8).Value = "Greatest Total Volume"
    
    ' Looping through every worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Skip 'Results' sheet
        If ws.Name <> "Results" Then
        
            ' Find last row of data
            last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            prev_quart = ""
            prev_ticker = ""
            
            ' Looping through data
            For i = 2 To last_row
            
                ' Get ticker
                curr_ticker = ws.Cells(i, 1).Value
                
                ' Get date
                Dim dateValue As String
                dateValue = CStr(ws.Cells(i, 2).Value)
                
                ' Get year
                Dim year As Integer
                year = CInt(Left(dateValue, 4))
                
                ' Get month
                Dim month As Integer
                month = CInt(Mid(dateValue, 5, 2))
                
                ' Calculate quarter
                Dim quarter As Integer
                quarter = WorksheetFunction.RoundUp(month / 3, 0)
                
                ' Combine as string
                curr_quart = year & "-Q" & quarter
                
                ' If ticker or quarter change, calculate values
                If curr_ticker <> prev_ticker Or curr_quart <> prev_quart Then
                    
                    ' If not first row
                    If prev_ticker <> "" Then

                        ' Get opening price for quarter
                        s_price = ws.Cells(quart_start, 3).Value
                        
                        ' Get closing price for quarter
                        e_price = ws.Cells(quart_end, 6).Value
                        
                        ' Calculate volume for quarter
                        volume = Application.Sum(ws.Range(ws.Cells(quart_start, 7), ws.Cells(quart_end, 7))) ' Assumes volume is in Column G
                        
                        ' Calculate percentage change for quarter
                        If s_price <> 0 Then
                            per_change = ((e_price - s_price) / s_price) * 100
                        Else
                            per_change = 0 ' Avoid division by zero
                        End If

                        ' Print quarterly data to 'Results' sheet
                        With res
                            Dim resultRow As Long
                            resultRow = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Row
                            .Cells(resultRow, 1).Value = prev_ticker ' A: Ticker
                            .Cells(resultRow, 2).Value = prev_quart ' B: Quarter
                            .Cells(resultRow, 3).Value = e_price - s_price ' C: Quarterly Change
                            .Cells(resultRow, 4).Value = Format(per_change, "0.00") & "%" ' D: Percentage Change
                            .Cells(resultRow, 5).Value = volume ' E: Total Stock Volume

                            ' Colouring
                            If e_price - s_price > 0 Then
                                .Cells(resultRow, 3).Interior.Color = RGB(144, 238, 144) ' Green for positive
                            ElseIf e_price - s_price < 0 Then
                                .Cells(resultRow, 3).Interior.Color = RGB(255, 182, 193) ' Red for negative
                            End If
                        End With
                        
                        ' Track overall totals for ticker
                        If os_price = 0 Then os_price = s_price
                        oe_price = e_price
                        volumeForTicker = volumeForTicker + volume

                        ' Track greatest quarterly increase, decrease, and volume
                        If per_change > great_incr Then
                            great_incr = per_change
                            great_incr_ticker = prev_ticker
                        End If

                        If per_change < great_decr Then
                            great_decr = per_change
                            great_decr_ticker = prev_ticker
                        End If

                        If volume > great_vol Then
                            great_vol = volume
                            great_vol_ticker = prev_ticker
                        End If
                    End If
                    
                    ' Reset variables for new quarter and ticker
                    prev_ticker = curr_ticker
                    prev_quart = curr_quart
                    quart_start = i
                    volumeForTicker = 0
                End If
                
                ' Update row where quarter ends
                quart_end = i
            Next i
            
            ' Handle last ticker's totals
            If prev_ticker <> "" Then
                ' Get opening price for quarter
                s_price = ws.Cells(quart_start, 3).Value
                        
                ' Get closing price for quarter
                e_price = ws.Cells(quart_end, 6).Value
                
                ' Calculate volume for quarter
                volume = Application.Sum(ws.Range(ws.Cells(quart_start, 7), ws.Cells(quart_end, 7)))
                
                ' Calculate percentage change for quarter
                If s_price <> 0 Then
                    per_change = ((e_price - s_price) / s_price) * 100
                Else
                    per_change = 0
                End If
                
                ' Print last ticker's quarterly data to 'Results' sheet
                With res
                    resultRow = .Cells(.Rows.Count, 1).End(xlUp).Offset(1, 0).Row
                    .Cells(resultRow, 1).Value = prev_ticker
                    .Cells(resultRow, 2).Value = prev_quart
                    .Cells(resultRow, 3).Value = e_price - s_price
                    .Cells(resultRow, 4).Value = Format(per_change, "0.00") & "%"
                    .Cells(resultRow, 5).Value = volume

                    ' Colouring for final ticker
                    If e_price - s_price > 0 Then
                        .Cells(resultRow, 3).Interior.Color = RGB(144, 238, 144)
                    ElseIf e_price - s_price < 0 Then
                        .Cells(resultRow, 3).Interior.Color = RGB(255, 182, 193)
                    End If
                End With

                ' Track overall totals for last ticker
                If o_per_change = 0 Then o_per_change = per_change
                
                ' Update greatest values for last ticker
                If o_per_change > great_incr Then
                    great_incr = o_per_change
                    great_incr_ticker = prev_ticker
                End If
                
                If o_per_change < great_decr Then
                    great_decr = o_per_change
                    great_decr_ticker = prev_ticker
                End If
                
                If volume > great_vol Then
                    great_vol = volume
                    great_vol_ticker = prev_ticker
                End If
            End If
        End If
    Next ws
    
    ' Print "greatests"
    With res
        .Cells(2, 9).Value = great_incr_ticker
        .Cells(2, 10).Value = Format(great_incr, "0.00") & "%"

        .Cells(3, 9).Value = great_decr_ticker
        .Cells(3, 10).Value = Format(great_decr, "0.00") & "%"

        .Cells(4, 9).Value = great_vol_ticker
        .Cells(4, 10).Value = great_vol
    End With

End Sub
