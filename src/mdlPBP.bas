Attribute VB_Name = "mdlPBP"
'author: Oleg Pyatakov
'email: oleg@pyatakov.com

Const maxPeriods = 360

Option Explicit

Function fPBP(varCashFlowRange As Range, Optional varDiscountRate As Variant, Optional varGrowthRate As Variant) As Variant

If IsMissing(varDiscountRate) Then varDiscountRate = 0
If IsMissing(varGrowthRate) Then varGrowthRate = 0

Dim varNPV As Double
Dim I, Periods As Integer
Dim varCF() As Double
Dim varDCF() As Double
Dim varPV() As Double

varNPV = varCashFlowRange(1, 1)

If varGrowthRate <> 0 Then
    Periods = WorksheetFunction.Max(maxPeriods, varCashFlowRange.Columns.Count)
Else
    Periods = varCashFlowRange.Columns.Count
End If

ReDim varCF(1 To Periods)
ReDim varDCF(1 To Periods)
ReDim varPV(1 To Periods)

For I = 1 To Periods
    If I <= varCashFlowRange.Columns.Count Then
        varCF(I) = varCashFlowRange(1, I)
    Else
        varCF(I) = varCF(I - 1) * varGrowthRate
    End If
    
    varDCF(I) = varCF(I) / (1 + varDiscountRate) ^ (I - 1)
    If I = 1 Then
        varPV(I) = varDCF(I)
    Else
        varPV(I) = varPV(I - 1) + varDCF(I)
    End If
Next

If varPV(Periods) < 0 Then
    'fPBP = ">" & Format(Periods - 1, "0")
    fPBP = "-"
    Exit Function
End If

For I = (Periods - 1) To 1 Step -1
    If varPV(I) < 0 Then
        fPBP = I - 1 - varPV(I) / varDCF(I + 1)
        Exit Function
    End If
Next

fPBP = "-"

End Function
