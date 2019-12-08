Attribute VB_Name = "mdlCAGR"
'author: Oleg Pyatakov
'email: oleg@pyatakov.com

Option Explicit

Function fCAGR(rValuesRange As Range) As Double

Dim dEndValue As Double
Dim iPeriods As Integer

If rValuesRange.Rows.Count > rValuesRange.Columns.Count Then
    dEndValue = rValuesRange(rValuesRange.Rows.Count, 1).Value
    iPeriods = rValuesRange.Rows.Count - 1
Else
    dEndValue = rValuesRange(1, rValuesRange.Columns.Count).Value
    iPeriods = rValuesRange.Columns.Count - 1
End If

fCAGR = (dEndValue / rValuesRange(1, 1).Value) ^ (1 / iPeriods) - 1

End Function

Function fCAGR2(dStartValue As Double, dEndValue As Double, dStartYear As Double, dEndYear As Double) As Double

fCAGR2 = (dEndValue / dStartValue) ^ (1 / (dEndYear - dStartYear)) - 1

End Function
