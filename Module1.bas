Attribute VB_Name = "Module1"
Sub main()
Dim ws As Worksheet
For Each ws In Worksheets
'hold variables for ticker, volume, open, close, and column placement,and other'
    Dim ticker As String
    Dim tsv As Double
    Dim table As Integer
    Dim lrow As Long
    Dim open_v As Double
    Dim close_v As Double
    Dim ticker_count As Long
    Dim percent_change As Double
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    table = 2
    tsv = 0
    ticker_count = 2
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'for loop, last row, utilize for 4 varaibles'
For C = 2 To lrow
'determine whether to continue or begin a new tick'
'everything below if is in relation to not being the same tick'
    If ws.Cells(C + 1, 1).Value <> ws.Cells(C, 1).Value Then
'extract new ticker string, related to "c" row'
        ticker = ws.Cells(C, 1).Value
'extract volume from new tick'
        tsv = tsv + ws.Cells(C, 7).Value
'extract open/close values'
        open_v = ws.Range("C" & ticker_count).Value
        close_v = ws.Range("F" & C).Value
        yearly_change = close_v - open_v
'extracting value for percent change, can't divide by 0, use if statement'
            If open_v = 0 Then
                percent_change = 0
            Else
                percent_change = yearly_change / open_v
            End If
'placement of values in column'
        ws.Range("I" & table).Value = ticker
        ws.Range("L" & table).Value = tsv
        ws.Range("J" & table).Value = yearly_change
        ws.Range("K" & table).Value = percent_change
        ws.Range("K" & table).NumberFormat = "0.00%"
        If ws.Range("J" & table).Value > 0 Then
            ws.Range("J" & table).Interior.Color = RGB(0, 225, 0)
         Else
            ws.Range("J" & table).Interior.Color = RGB(225, 0, 0)
        End If
'go to next row for each itteration'
        table = table + 1
        ticker_count = C + 1
'reset volume'
        tsv = 0
'finalize volume of ticks'
    Else
        tsv = tsv + ws.Cells(C, 7).Value
    End If
Next C
Next ws
End Sub


