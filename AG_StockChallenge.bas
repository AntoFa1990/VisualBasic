Attribute VB_Name = "Module1"
Sub Stock_Tickers()
'Declare and set worksheet
Dim ws As Worksheet

'Loop through all stocks for each year
For Each ws In Worksheets

'Create the column headings
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
Dim Ticker As String
Ticker = " "
Dim Ticker_volume As Double
Ticker_volume = 0

'Create variable to hold stock volume
Dim stock_volume As Double
stock_volume = 0
'Set initial and last row for worksheet
Dim Lastrow As Long
Dim i As Long
Dim j As Integer

'Define Lastrow of worksheet
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set new variables for prices and percent changes
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0
Dim price_change As Double
price_change = (close_price - open_price)
Dim price_change_percent As Double
price_change_percent = 0

Dim TickerRow As Long
TickerRow = 1
open_price = ws.Cells(2, 3).Value

'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow

'Ticker symbol output

stock_volume = stock_volume + ws.Cells(i, 7).Value

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
close_price = ws.Cells(i, 6).Value

TickerRow = TickerRow + 1
Ticker = ws.Cells(i, 1).Value
ws.Cells(TickerRow, 9).Value = Ticker
ws.Cells(TickerRow, 10).Value = close_price - open_price

'Calculate change in Price
If open_price <> 0 Then
price_change_percent = (close_price - open_price)

ws.Cells(TickerRow, 11).Value = (price_change_percent / open_price) * 100
ws.Cells(TickerRow, 11).NumberFormat = "0.0%"

End If

ws.Cells(TickerRow, 12).Value = stock_volume
ws.Cells(TickerRow, 12).NumberFormat = "#,##0"

stock_volume = 0

open_price = ws.Cells(i + 1, 3).Value

End If

Next i
'Declare variables for greatest volume, greatest increase and greatest decrease

Dim greatest_total_volume, GTV_ticker, greatest_increase, greatest_decrease As Integer


'Do loop of current worksheet to Lastrow
For i = 2 To Lastrow

'if next stocks total volume is greater than current stock, greatest total volume is equal to next stock
If ws.Cells(i + 1, 12).Value > ws.Cells(i, 12).Value Then
greatest_total_volume = ws.Cells(i + 1, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i + 1, 9).Value
ws.Cells(4, 17).Value = greatest_total_volume
ws.Cells(4, 17).NumberFormat = "#,##0"


End If

Next i
Dim xmax As Double
    Dim xmin As Double
    Dim TableRow As Integer

    For i = 2 To Lastrow
        If ws.Cells(i, 11).Value < Cells(i + 1, 11).Value Then
            xmin = ws.Cells(i, 11).Value
            ws.Cells(3, 17).Value = xmin
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).NumberFormat = "0.0%"

        End If

        If ws.Cells(i, 11).Value > ws.Cells(i + 1, 11).Value Then
            ws.Cells(2, 17).Value = xmax
            ws.Cells(2, 17).NumberFormat = "0.0%"
            ws.Cells(2, 16).Value = ws.Cells(i, 9)
            

        End If
    Next i
Next ws


End Sub
Sub Format_To_NumComma()
Attribute Format_To_NumComma.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Format_To_NumComma Macro
'

'
    Columns("L:L").Select
    Selection.NumberFormat = "#,##0"
End Sub
