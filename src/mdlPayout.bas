Attribute VB_Name = "mdlPayout"
'author: Oleg Pyatakov
'email: oleg@pyatakov.com

Option Explicit

Function fPayout(dPayoutBase As Double, rFutureCashFlow As Range, Optional dWACC As Double, Optional dMinCashLimit As Double) As Double

If IsMissing(dWACC) Then dWACC = 0
If IsMissing(dMinCashLimit) Then dMinCashLimit = 0

Dim I As Integer
Dim dCash As Double
Dim dMinCashValueTemp As Double
Dim dMinCashValueLimit As Double

dMinCashValueLimit = 0
dCash = 0

For I = 1 To rFutureCashFlow.Columns.Count
    dCash = dCash * (1 + dWACC) + rFutureCashFlow(1, I).Value
    dMinCashValueTemp = -(dCash - dMinCashLimit) / ((1 + dWACC) ^ I)
    If dMinCashValueTemp > dMinCashValueLimit Then
        dMinCashValueLimit = dMinCashValueTemp
    End If
Next

fPayout = WorksheetFunction.Max(dPayoutBase - dMinCashValueLimit, 0)

End Function
