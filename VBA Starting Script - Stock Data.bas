Attribute VB_Name = "Module1"
Sub TickerData()

'define variables
Dim tickerArray() As Variant
Dim Rg As Range

'Set arrays
tickerArray = Range("A1:A70926")

'Print Values
Range("I1:I70926") = tickerArray

'Set Range

Set Rg = Range("I1:I70926").CurrentRegion

'Remove Duplicates

Rg.RemoveDuplicates Columns:=1, Header:=xlYes

'Erase Ticker

Dim Full_Ticker_Array() As Variant
Full_Ticker_Array = Range("A1:A70926")
End Sub

Sub TickerLoop()
'define variables
Dim TickerSymbol As String
Dim TickerTotal As Long
Dim lastrow As Long
Dim Yearlychange As Long
Dim PercentageChanged As Long
Dim OpenPrice As Long
Dim ClosePrice As Long
Dim close_open As Long


'Print summary table
Cells(1, 11).Value = "Percentage_Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(3, 14).Value = "Greatest % Increase"
Cells(4, 14).Value = "Greatest % Decrease"
Cells(5, 14).Value = "Greatest Total Volume"

'calculating
Yearlychange = ClosePrice - OpenPrice
PercentageChange = close_open - OpenPrice

TickerTotal = 0
Dim summary_table_row As Long
summary_table_row = 2

'For Loop
For i = 2 To 70927
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    TickerSymbol = Cells(i, 1).Value
    OpenPrice = Cells(i + 1, 3).Value
    
    ClosePrice = Cells(i, 6).Value
    TickerTotal = TickerTotal + Cells(i, 7).Value
    
    close_open = ClosePrice - OpenPrice
    PercentageChanged = close_open / ClosePrice
ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
OpenPrice = Cells(i, 3).Value

Range("I" & summary_table_row).Value = TickerSymbol
Range("L" & summary_table_row).Value = TickerTotal
Range("J" & summary_table_row).Value = close_open
Range("K" & summary_table_row).Value = PercentageChange
Range("N" & summary_table_row).Value = OpenPrice
Range("O" & summary_table_row).Value = ClosePrice
summary_table_row = summary_table_row + 1

TickerTotal = 0
End If

Next i

End Sub
