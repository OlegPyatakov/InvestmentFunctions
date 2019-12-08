Attribute VB_Name = "mdlIncomeTax"
'author: Oleg Pyatakov
'email: oleg@pyatakov.com

Option Explicit

Function fIncomeTax(rProfitRange As Range, dTaxRate As Double, Optional dPriorProfitsAndLosses As Double) As Double

If rProfitRange(1, rProfitRange.Columns.Count).Value <= 0 Then
    fIncomeTax = 0
    Exit Function
End If

If IsMissing(dPriorProfitsAndLosses) Then
    dPriorProfitsAndLosses = 0
End If

Dim I As Integer

For I = 1 To rProfitRange.Columns.Count
    If dPriorProfitsAndLosses > 0 Then
        dPriorProfitsAndLosses = 0
    End If
    dPriorProfitsAndLosses = dPriorProfitsAndLosses + rProfitRange(1, I).Value
Next

If dPriorProfitsAndLosses > 0 Then
    fIncomeTax = dPriorProfitsAndLosses * dTaxRate
End If

End Function
