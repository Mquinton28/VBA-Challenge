Attribute VB_Name = "Module1"
Sub TickerTotal()
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
Erase tickerArray()

Dim Full_Ticker_Array() As Variant
Full_Ticker_Array = Range("A1:A70926")

End Sub
Sub StockData()
'Define Variables
Dim i As Integer
Dim j As Integer
Dim lastrow As Long
Dim Ticker As String
Dim OpenDate As Double
Dim CloseDate As Double
Dim YearlyChanged As Double
Dim PercentageChanges As Double
Dim TotalStockVol As Long

'Locate Variables
Ticker = Cells(2, 1).Value
lastrow = Range("A1").End(xlDown).Row
OpenDate = Range("C1").End(xlDown).Row
CloseDate = Range("F1").End(xlDown).Row

'Store summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'For Loop
For i = 2 To 70926
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value

    
    Range("I" & Summary_Table_Row).Value = Ticker
    Range("
    
    End If


Next i
End Sub
