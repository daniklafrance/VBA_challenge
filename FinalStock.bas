Attribute VB_Name = "Module1"
Sub stock()

For Each ws In Worksheets

'set headers
ws.Cells(1, 9) = "ticker"
ws.Cells(1, 10) = "yearly change"
ws.Cells(1, 11) = "percent change"
ws.Cells(1, 12) = "total stock volume"
ws.Cells(2, 15) = "greatest % increase"
ws.Cells(3, 15) = "greatest % decrease"
ws.Cells(4, 15) = "greatest total volume"
ws.Cells(1, 16) = "ticker"
ws.Cells(1, 17) = "value"

'set variable for ticker info
Dim ticker As String

'set variable for total stock
Dim total_stock As Double
total_stock = 0

'set variable for the summary table
Dim summary As Integer
summary = 2

'set variables for the opening/closing/change stock info
Dim open_price As Double
open_price = ws.Cells(2, 3).Value
Dim closing_price As Double
Dim yearly_change As Double

'set variable for percent change
Dim percent_change As Double

'set range to end at last row
Dim last_row As Long
last_row = ws.Range("A" & Rows.Count).End(xlUp).Row

'loop through all rows
For i = 2 To last_row

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'find ticker info and put in summary table
    ticker = ws.Cells(i, 1).Value
    ws.Range("I" & summary).Value = ticker
    
    'input total value into the summary table
    total_stock = total_stock + ws.Cells(i, 7).Value
    ws.Range("L" & summary).Value = total_stock
    
    'complete calculation for the yearly change in the stock
    close_price = ws.Cells(i, 6).Value
    yearly_change = close_price - open_price
    ws.Range("J" & summary).Value = yearly_change
    
    'complete the calculation for the percent change
    percent_change = yearly_change / open_price
    ws.Range("K" & summary).Value = percent_change
    ws.Range("K" & summary).NumberFormat = "0.00%"
    
    'move to the next summary box and reset the count
    summary = summary + 1
    total_stock = 0
    open_price = ws.Cells(i + 1, 3)

    Else
    
    'add value to the total stock for the next ticker symbol
    total_stock = total_stock + ws.Cells(i, 7).Value

    End If
    
Next i

'set range to end at last row of summary
Dim last_row2 As Long
last_row2 = ws.Range("J" & Rows.Count).End(xlUp).Row

For i = 2 To last_row2

'colour fill cells based on positive vs negative growth
    If ws.Cells(i, 10).Value > 0 Then

    ws.Cells(i, 10).Interior.ColorIndex = 4

    Else

    ws.Cells(i, 10).Interior.ColorIndex = 3

    End If

'change column values
ws.Cells(i, 10).NumberFormat = "0.00"

Next i

'find the max/min results for this table
Dim greatinc As Double
Dim greatdec As Double
Dim totalvol As Double
Dim ticker2 As String

greatinc = Application.WorksheetFunction.Max(ws.Range("K:K"))
ws.Range("Q2").Value = greatinc
ws.Range("Q2").NumberFormat = "0.00%"

greatdec = Application.WorksheetFunction.Min(ws.Range("K:K"))
ws.Range("Q3").Value = greatdec
ws.Range("Q3").NumberFormat = "0.00%"

totalvol = Application.WorksheetFunction.Max(ws.Range("L:L"))
ws.Range("Q4").Value = totalvol
ws.Range("Q4").NumberFormat = "0"

'autofit each column of each worksheet
ws.Range("A:Q").EntireColumn.autofit

Next ws
    
End Sub
