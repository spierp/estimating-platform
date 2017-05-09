Attribute VB_Name = "aGeneralMaster"
Sub aaGeneralSheets()
    If range("coverpage").Value = "Yes" Then
        Sheets("cover").Visible = True
        Call coverPage
    Else
        If Sheets("cover").Visible = True Then
            Sheets("cover").Visible = False
        End If
    End If
    If range("tablecontents").Value = "Yes" Then
        Sheets("TOC").Visible = True
'        Call tableContents
    Else
        If Sheets("TOC").Visible = True Then
            Sheets("TOC").Visible = False
        End If
    End If
    If range("notesquals").Value = "Yes" Then
        Call notesQualsCopy
        Call notesQualsInsert
        Call notesQualsFormat
    Else
        If Sheets("N+Q").Visible = True Then
            Sheets("N+Q").Visible = False
        End If
    End If
    If range("bim").Value = "Yes" Then
'        Call bim
    Else
        If Sheets("BIM").Visible = True Then
            Sheets("BIM").Visible = False
        End If
    End If
        
    Worksheets("dashboard").Activate
End Sub

