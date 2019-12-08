Attribute VB_Name = "mdlPI"
'author: Oleg Pyatakov
'email: oleg@pyatakov.com

Option Explicit

Function fPI(varCashFlowRange As Range, Optional varDiscountRate As Variant) As Variant

If IsMissing(varDiscountRate) Then
    varDiscountRate = 0
End If

Dim dCurrDiscount As Double
Dim iPeriod, iPeriods As Integer
Dim dCurrPV, dCurrPosPV, dCurrNegPV As Double

dCurrPosPV = 0
dCurrNegPV = 0
dCurrDiscount = 1

For iPeriod = 1 To varCashFlowRange.Columns.Count
    dCurrDiscount = dCurrDiscount * (1 + varDiscountRate)
    dCurrPV = varCashFlowRange(1, iPeriod) / dCurrDiscount
    If dCurrPV > 0 Then
        dCurrPosPV = dCurrPosPV + dCurrPV
    ElseIf dCurrPV < 0 Then
        dCurrNegPV = dCurrNegPV + dCurrPV
    End If
Next

fPI = -dCurrPosPV / dCurrNegPV

End Function

