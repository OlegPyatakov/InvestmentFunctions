VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Macro for Excel to automatically update links to this addin
'Source: http://www.planetaexcel.ru/forum/index.php?PAGE_NAME=message&FID=1&TID=56510&TITLE_SEO=56510&MID=471639&tags=&q=%D0%BD%D0%B0%D0%B4%D1%81%D1%82%D1%80%D0%BE%D0%B9%D0%BA%D0%B0+%D0%B4%D0%BB%D1%8F+%D0%BD%D0%B5%D1%81%D0%BA%D0%BE%D0%BB%D1%8C%D0%BA%D0%B8%D1%85+%D0%BF%D0%BE%D0%BB%D1%8C%D0%B7%D0%BE%D0%B2%D0%B0%D1%82%D0%B5%D0%BB%D0%B5%D0%B9&DATE_CHANGE=0&order=relevance&s=%D0%9D%D0%B0%D0%B9%D1%82%D0%B8#message471639

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub Workbook_Open()
    Set App = Application
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Dim sMulTEx As String, sLnk, wsSh As Worksheet
    
    'Need to work only if this file is stored as addin
    If Not UCase(ThisWorkbook.Name) Like "*" & ".XLAM" Then
        Exit Sub
    End If
    
    sMulTEx = UCase(ThisWorkbook.Name)
    On Error GoTo exit_
    With Wb
        For Each sLnk In .LinkSources(Type:=xlExcelLinks)
            If UCase(sLnk) Like "*" & sMulTEx Then
                If UCase(sLnk) <> sMulTEx Then
                    .ChangeLink Name:=sLnk, NewName:=sMulTEx
                    For Each wsSh In .Worksheets
                        wsSh.Calculate
                    Next
                End If
                Exit For
            End If
        Next
    End With
exit_:
End Sub


