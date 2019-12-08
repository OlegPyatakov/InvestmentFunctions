Attribute VB_Name = "mdlLimit"
'author: Oleg Pyatakov
'email: oleg@pyatakov.com

Option Explicit

Function fLimit(varInput, varLimit1, varLimit2)

fLimit = WorksheetFunction.Max(varInput, WorksheetFunction.Min(varLimit1, varLimit2))
fLimit = WorksheetFunction.Min(fLimit, WorksheetFunction.Max(varLimit1, varLimit2))

End Function
