Attribute VB_Name = "Module1"
Sub stocktime()

'I need to make sure this applies to each worksheet-------
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

'My column headers-------
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percentage Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"

'Declare-------
    Dim ticker_name As String
    Dim yearly_change, percent_change, stock_volume, first_open, last_close As Double
    Dim summary_table_row As Integer

'Variables-------
    stock_volume = 0
'Variables - Sum table row
    summary_table_row = 2
'Variables - open price
    first_open = Cells(2, 3).Value
'Variables - Last row
    last_row = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Move ticker symbols over from 1 to 9, this will need to look like the credit card code problem where we recognize
'a different (<>) ticker along with its total and move it to another column-------
        For i = 2 To last_row

    'Checking for same ticker type, it not, then it goes on to the next-------
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ticker_name = Cells(i, 1).Value
                stock_volume = stock_volume + Cells(i, 7).Value
            'Apply ticker and stock volume to ranges I and L along with sum table row
                Range("I" & summary_table_row).Value = ticker_name
                Range("L" & summary_table_row).Value = stock_volume
            'Last close is going to be the last cell of column 6 within this loop
                last_close = Cells(i, 6).Value
            'Yearly change will be the last cell of column 6 within this loop minus the first, opening price within this loop
                yearly_change = (last_close - first_open)
            'Apply Yearly Change to range J along with the next sum table tow
                Range("J" & summary_table_row).Value = yearly_change
            
                'I need to override not being able to divide by 0. What i'll need to do is if the opening price is 0 then
                'the percent change will need to be 0
                'When it is not 0, the percent change will simply be the yearly change divided by the original amount(i.e., first_open)
                    If first_open = 0 Then
                        percent_change = 0
                
                    Else
                        percent_change = yearly_change / first_open
                    End If
            'Once I have the percent change I need to apply them to range K and turn it into a percentage, I don't know how to
            'do this within one line so I'll try two...
                Range("K" & summary_table_row).Value = percent_change
                Range("K" & summary_table_row).NumberFormat = "0.00%"
            'Reset the stock volume
                stock_volume = 0
            'Reset the sum table and then make sure to add one
                summary_table_row = summary_table_row + 1
            'Reset my fopen price
                first_open = Cells(i + 1, 3)
            
            Else
                stock_volume = stock_volume + Cells(i, 7).Value
               
        'I will need a conditional within my forloop for color coding my yearly change. I think I should be fine with just less than 0
        'or > 0
                If Cells(i, 10).Value < 0 Then
                    Cells(i, 10).Interior.ColorIndex = 3
                Else
                    Cells(i, 10).Interior.ColorIndex = 4
                End If

            End If
    
    Next i
Next WS

End Sub

