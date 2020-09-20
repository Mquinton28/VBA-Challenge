Attribute VB_Name = "Module1"
Sub TickerTotal()
'define variables
Dim ticker() As String
Dim datenum() As Long
Dim opennum() As Double
Dim highnum() As Double
Dim lownum() As Double
Dim closenum() As Double
Dim volnum() As Long

'Set arrays
ticker = Range("A2:A70926").Value
datenum = Range("B2:B70926").Value
opennum = Range("C:C70926").Value
highnum = Range("D2:D70926").Value
lownum = Range("E2:E70926").Value
closenum = Range("F2:F70926").Value
volnum = Range("G2:G70926").Value

'Create For Loop
Dim Row As String
Dim Column As String

For Row = 2 To 70926
    For Column = 2 To 7
    Next Column
Next Row

End Sub
