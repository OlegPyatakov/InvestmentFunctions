Attribute VB_Name = "mdlExtend"
'author: Oleg Pyatakov
'email: oleg@pyatakov.com

Option Explicit

Function fExtendAP(vInput As Variant, iPeriods As Variant, Optional dGrowthRate As Variant) As Variant

If iPeriods = 0 Then
    fExtendAP = vInput
    Exit Function
End If

If IsMissing(dGrowthRate) Then dGrowthRate = 0

Dim dArray() As Double
Dim rCell As Range
Dim I As Integer

ReDim Preserve dArray(0)

If TypeName(vInput) = "Range" Then
    I = 0
    For Each rCell In vInput.Cells
        ReDim Preserve dArray(I)
        dArray(I) = CDbl(rCell.Value)
        I = I + 1
    Next
    I = UBound(dArray)
Else
    I = 0
    dArray(I) = CDbl(vInput)
End If

For I = I + 1 To I + iPeriods
    ReDim Preserve dArray(I)
    dArray(I) = CDbl(dArray(I - 1)) + dGrowthRate
Next

fExtendAP = dArray
End Function

Function fExtendGP(vInput As Variant, iPeriods As Variant, Optional dGrowthRate As Variant) As Variant

If iPeriods = 0 Then
    fExtendGP = vInput
    Exit Function
End If

If IsMissing(dGrowthRate) Then dGrowthRate = 1

Dim dArray() As Double
Dim rCell As Range
Dim I As Integer

ReDim Preserve dArray(0)

If TypeName(vInput) = "Range" Then
    I = 0
    For Each rCell In vInput.Cells
        ReDim Preserve dArray(I)
        dArray(I) = CDbl(rCell.Value)
        I = I + 1
    Next
    I = UBound(dArray)
Else
    I = 0
    dArray(I) = CDbl(vInput)
End If

For I = I + 1 To I + iPeriods
    ReDim Preserve dArray(I)
    dArray(I) = CDbl(dArray(I - 1)) * dGrowthRate
Next

fExtendGP = dArray
End Function
