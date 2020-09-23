Sub Stk_Mkt_Analysis()

'create a loop to go through all worksheets in the workbook for yearly stock market data
'challenge to add additional yearly analysis

'first variable for the worksheets

    Dim ws As Worksheet

'Start looping through the worksheets

    For Each ws In Worksheets

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

'second set of variables for columns 1 to 6

    Dim ticker As String

    Dim volume As Double
    volume = 0

    Dim rowcount As Long
    rowcount = 2

    Dim year_close As Double
    year_close = 0

    Dim year_change As Double
    year_change = 0

    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastrow

'first conditional statment
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            Year_open = ws.Cells(i, 3).Value
        End If

        volume = volume + ws.Cells(i, 7)

'second conditional statement
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

            ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(rowcount, 12).Value = volume

            year_close = ws.Cells(i, 6).Value

            year_change = year_close - Year_open
            ws.Cells(rowcount, 10).Value = year_change

        If year_change >= 0 Then
            ws.Cells(rowcount, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(rowcount, 10).Interior.ColorIndex = 3
        End If

'third conditional statement for the percentages

        If Year_open = 0 And year_close = 0 Then

        Percent_change = 0

            ws.Cells(rowcount, 11).Value = Percent_change
            ws.Cells(rowcount, 11).NumberFormat = "0.00%"

        ElseIf Year_open = 0 Then

            Dim percent_change_new As String
            percent_change_new = "New Stock"
            ws.Cells(rowcount, 11) = Percent_change

        Else
    
            Percent_change = year_change / Year_open
            ws.Cells(rowcount, 11).Value = Percent_change
            ws.Cells(rowcount, 11).NumberFormat = "0.00%"
        End If

'add one to the rowcount then move to the next empty row

        rowcount = rowcount + 1

'reset factors being analyized

        volume = 0
        Year_open = 0
        year_close = 0
        year_change = 0
        Percent_change = 0
        
        End If

    Next i

'final challenge request for yearly analysis

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

'Adding challenge variables to yearly analysis

        Dim best_stk As String
        Dim best_value As Double

        best_value = ws.Cells(2, 11).Value

        Dim wrst_stk As String
        Dim wrst_value As Double

        wrst_value = ws.Cells(2, 11).Value

        Dim highest_stk As String
        Dim highest_value As Double

        highest_value = ws.Cells(2, 12).Value

'redefine lastrow for challenge as first lastrow request was not working

        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        For j = 2 To lastrow

'conditional for challenge request

        If ws.Cells(j, 11).Value > best_value Then
            best_value = ws.Cells(j, 11).Value
            best_stk = ws.Cells(j, 9).Value

        End If

        If ws.Cells(j, 11).Value < wrst_value Then
            wrst_value = ws.Cells(j, 11).Value
            wrst_stk = ws.Cells(j, 9).Value
        End If

        If ws.Cells(j, 12).Value > highest_value Then
            highest_value = ws.Cells(j, 12).Value
            highest_stk = ws.Cells(j, 9).Value
        End If

    Next j

'fill summary table for yearly analysis challenge

        ws.Range("P2").Value = best_stk
        ws.Range("Q2").Value = best_value
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("P3").Value = wrst_stk
        ws.Range("Q3").Value = wrst_value
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("P4").Value = highest_stk
        ws.Range("Q4").Value = highest_value

    Next ws

End Sub

