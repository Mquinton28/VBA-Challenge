Attribute VB_Name = "Module1"
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
For i = 2 To lastrow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value

    
    Range("I" & Summary_Table_Row).Value = Ticker
    
    End If


Next i
End Sub
