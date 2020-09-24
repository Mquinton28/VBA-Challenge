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
Dim StockVol As Double
Dim Yearlychange As Double
Dim PercentageChange As Double
Dim OpenPrice As Double
OpenPrice = Cells(2, 3).Value
Dim ClosePrice As Double


'Print summary table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage_Change"
Cells(1, 12).Value = "Total Stock Volume"
'Cells(2, 14).Value = "Greatest % Increase"
'Cells(3, 14).Value = "Greatest % Decrease"
'Cells(4, 14).Value = "Greatest Total Volume"

'calculating
'Yearlychange = ClosePrice - OpenPrice
'PercentageChange = close_open - OpenPrice

StockVol = 0
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'For Loop
For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    TickerSymbol = Cells(i, 1).Value
    
    
    ClosePrice = Cells(i, 6).Value
    StockVol = StockVol + Cells(i, 7).Value
    
    Yearlychange = ClosePrice - OpenPrice
    PercentageChange = Yearlychange / OpenPrice
    PercentageChange = Round(PercentageChange, 2)

Range("L" & Summary_Table_Row).Value = StockVol
Range("J" & Summary_Table_Row).Value = Yearlychange
Range("K" & Summary_Table_Row).Value = PercentageChange
Summary_Table_Row = Summary_Table_Row + 1
StockVol = 0
    
    OpenPrice = Cells(i + 1, 3).Value
    

ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then

StockVol = StockVol + Cells(i, 7).Value

End If

Next i

End Sub

