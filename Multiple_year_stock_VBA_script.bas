Attribute VB_Name = "Module1"
Sub stock_checker()

    ' Initialize worksheet variables
    Dim ws As Worksheet
    Dim WS_Count As Integer
    Dim sheet As Integer
    
    ' Initialize calculation variables
    Dim ticker_label As String
    Dim ticker_count As Integer
    Dim start_row As Integer
    Dim last_row As Long
    Dim earliest_data As Date
    Dim latest_date As Date
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_sum As Double
    
    ' Count worksheets and loop through each worksheet
    WS_Count = ActiveWorkbook.Worksheets.Count
    For sheet = 1 To WS_Count
        Worksheets(sheet).Activate
        Set ws = ThisWorkbook.Sheets(sheet)
           
        ' Check for the last row of data
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set initial variable values
        ticker_label = ""
        ticker_count = 0
        start_row = 2
        open_price = 0
        close_price = 0
        stock_sum = 0
           
        ' Loop through full data set
            For r = 2 To last_row
                ' Calculate total stock sum
                stock_sum = stock_sum + (Cells(r, 7).Value)
                
                ' Set column headings if not yet existing
                If Range("I1") = "" Then
                    Range("I" & start_row - 1).Value = "Ticker"
                    Range("J" & start_row - 1).Value = "Yearly Change"
                    Range("K" & start_row - 1).Value = "Percent Change"
                    Range("L" & start_row - 1).Value = "Total Stock Volume"
                    Range("P" & start_row - 1).Value = "Ticker"
                    Range("Q" & start_row - 1).Value = "Value"
                    Range("O" & start_row).Value = "Greatest % Increase"
                    Range("O" & start_row + 1).Value = "Greatest % Decrease"
                    Range("O" & start_row + 2).Value = "Greatest Total Volume"
                    open_price = CDbl(Cells(r, 3).Value)
                End If
                ' Check each new row of data for changes in ticker label
                If Cells(r, 1).Value <> Cells(r + 1, 1).Value Then
                    '  Calculate results of ticker grouping and input values
                    ticker_label = Cells(r, 1).Value
                    Range("I" & start_row).Value = ticker_label
                    close_price = CDbl(Cells(r, 6).Value)
                    yearly_change = close_price - open_price
                    Range("J" & start_row).Value = yearly_change
                    percent_change = yearly_change / open_price
                    Range("K" & start_row).Value = percent_change
                    Range("K" & start_row).NumberFormat = "0.00%"
                    Range("L" & start_row).Value = stock_sum
                    open_price = CDbl(Cells(r + 1, 3).Value)
                    ' Set input row below previous value
                    start_row = start_row + 1
                    ' Set stock sum variable back to 0
                    stock_sum = 0
                End If
                
            Next r
        
            ' Initialize variables for greatest percentage increase, decrease and stock volume
            Dim greatest_incr As Double
            Dim greatest_decr As Double
            Dim greatest_vol As Double
            Dim I As Integer
            ' Set initial summary variable values
            greatest_incr = 0
            greatest_decr = 0
            greatest_vol = 0
            
            
            'Loop through results to find greatest increase and decrease in percentgae change and greatest stock volume
            For I = 2 To (start_row - 1)
                 If Cells(I, 11).Value > greatest_incr Then
                        greatest_incr = Cells(I, 11).Value
                        Range("Q2").Value = greatest_incr
                        Range("Q2").NumberFormat = "0.00%"
                        Range("P2").Value = Cells(I, 9).Value
                End If
                If Cells(I, 11).Value < greatest_decr Then
                        greatest_decr = Cells(I, 11).Value
                        Range("Q3").Value = greatest_decr
                        Range("Q3").NumberFormat = "0.00%"
                        Range("P3").Value = Cells(I, 9).Value
                End If
                If Cells(I, 12).Value > greatest_vol Then
                        greatest_vol = Cells(I, 12).Value
                        Range("Q4").Value = greatest_vol
                        Range("P4").Value = Cells(I, 9).Value
                End If
            Next I
        
            ' Apply autofit column width to all result cells
            Columns("I:Q").Select
            Columns("I:Q").EntireColumn.AutoFit
        
            ' Apply conditional formatting for values > 0 (Green)
            With ws.Range("J2:K" & start_row).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
                .Interior.Color = RGB(0, 255, 0) ' Green
            End With
            
            ' Apply conditional formatting for values < 0 (Red)
            With ws.Range("J2:K" & start_row).FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
                .Interior.Color = RGB(255, 0, 0) ' Red
            End With
               
    Next sheet
    
End Sub


Sub Clear_results()
' Clears results of stock checker function
    
WS_Count = ActiveWorkbook.Worksheets.Count
    
    For sheet = 1 To WS_Count
        Worksheets(sheet).Activate
    
        Columns("I:Q").Select
        Selection.ClearContents
        
    Next sheet
End Sub
