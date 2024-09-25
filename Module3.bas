Attribute VB_Name = "Module3"
Option Explicit
Sub weekendgrey()
Dim ws As Worksheet
Set ws = ActiveSheet
Dim nc As Integer
Dim value1 As Integer
Dim value2 As Variant

nc = WorksheetFunction.CountA(Range("1:1")) - 3
Dim i As Integer
For i = 4 To nc
value2 = Cells(1, i).value
value1 = WorksheetFunction.Weekday(value2)
If value1 = 1 Or value1 = 7 Then
Cells(1, i).EntireColumn.Interior.ColorIndex = 2

End If
Next i


End Sub
