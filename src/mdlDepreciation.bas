Attribute VB_Name = "mdlDepreciation"
'author: Oleg Pyatakov
'email: oleg@pyatakov.com

Option Explicit

Function fDepreciation(rCapex As Range, dDepreciationPeriod As Double) As Double

Dim rCell As Range
Dim iCurrCapexYear As Integer

fDepreciation = 0

Dim dDepreciationRate  As Double

For iCurrCapexYear = 1 To WorksheetFunction.Min(rCapex.Cells.Count, WorksheetFunction.RoundUp(dDepreciationPeriod, 0))
    If iCurrCapexYear <> dDepreciationPeriod And iCurrCapexYear = WorksheetFunction.RoundUp(dDepreciationPeriod, 0) Then
        dDepreciationRate = (dDepreciationPeriod - WorksheetFunction.RoundDown(dDepreciationPeriod, 0)) / dDepreciationPeriod
    Else
        dDepreciationRate = 1 / dDepreciationPeriod
    End If
    'Debug.Print iCurrCapexYear & " - " & dDepreciationRate
    fDepreciation = fDepreciation + rCapex.Cells(rCapex.Cells.Count - iCurrCapexYear + 1) * dDepreciationRate
Next

End Function
