Attribute VB_Name = "aGeneralMaster"
Sub aaGeneralSheets()

'GENERAL TAB
    Call progressIndicator_Begin("General Tabs")
    pb.AddCaption "Working on general tabs..."
    If Range("coverpage").Value = "Yes" Then
        pb.AddCaption "Formatting coverpage..."
        Sheets("cover").Visible = True
        Call coverPage
    Else
        If Sheets("cover").Visible = True Then
            Sheets("cover").Visible = False
        End If
    End If
    pb.AddProgress 10
    
    If Range("tablecontents").Value = "Yes" Then
        pb.AddCaption "Creating Table of Contents..."
        Sheets("TOC").Visible = True
        Call tableofContents
    Else
        If Sheets("TOC").Visible = True Then
            Sheets("TOC").Visible = False
        End If
    End If
    pb.AddProgress 20
    
    If Range("notesquals").Value = "Yes" Then
        pb.AddCaption "Scrubbing Notes & Quals data..."
        Sheets("N+Q").Visible = True
        Call notesQualsCopy
        pb.AddCaption "Copying Notes & Quals data..."
        Call notesQualsInsert
        pb.AddCaption "Formatting Notes & Quals data..."
        Call notesQualsFormat
    Else
        If Sheets("N+Q").Visible = True Then
            Sheets("N+Q").Visible = False
        End If
    End If
    pb.AddProgress 40
    
    Call BIM

    pb.AddProgress 20
    
    If Range("executive_summary").Value = "Yes" Then
        pb.AddCaption "Creating Executive Summary..."
        Sheets("execSum").Visible = True
        Call execparts
        Call execpage
    Else
        If Sheets("execSum").Visible = True Then
            Sheets("execSum").Visible = False
        End If
    End If
    
    pb.AddProgress 10
    
    Call progressIndicator_End
   
    Worksheets("dashboard").Activate
End Sub

Sub aaGeneralCover()
    Call progressIndicator_Begin("Cover Page")
    
    If Range("coverpage").Value = "Yes" Then
        pb.AddCaption "Formatting coverpage..."
        Sheets("cover").Visible = True
        Call coverPage
        pb.AddProgress 75
    Else
        If Sheets("cover").Visible = True Then
            Sheets("cover").Visible = False
        End If
    End If
    
    Call progressIndicator_End
    Worksheets("dashboard").Activate
End Sub

Sub aaGeneralTOC()
    Call progressIndicator_Begin("TOC")

    If Range("tablecontents").Value = "Yes" Then
        pb.AddCaption "Creating Table of Contents..."
        Sheets("TOC").Visible = True
        pb.AddProgress 10
        Call tableofContents
        pb.AddProgress 85
    Else
        If Sheets("TOC").Visible = True Then
            Sheets("TOC").Visible = False
        End If
    End If

    Call progressIndicator_End
    Worksheets("dashboard").Activate
End Sub

Sub aaGeneralNQ()
    Call progressIndicator_Begin("Notes and Qualifications")
    pb.AddProgress 10
    If Range("notesquals").Value = "Yes" Then
        pb.AddCaption "Scrubbing Notes & Quals data..."
        Sheets("N+Q").Visible = True
        Call notesQualsCopy
        pb.AddProgress 25
        pb.AddCaption "Copying Notes & Quals data..."
        Call notesQualsInsert
        pb.AddProgress 25
        pb.AddCaption "Formatting Notes & Quals data..."
        Call notesQualsFormat
        pb.AddProgress 25
    Else
        If Sheets("N+Q").Visible = True Then
            Sheets("N+Q").Visible = False
        End If
    End If

    Call progressIndicator_End
    Worksheets("dashboard").Activate
End Sub

Sub aaGeneralBIM()
    Call progressIndicator_Begin("BIM Supplement")
    pb.AddProgress 10
    pb.AddCaption "Working on BIM supplemental tabs..."
    pb.AddProgress 25
    Call BIM
    pb.AddProgress 50
    Call progressIndicator_End
    Worksheets("dashboard").Activate
End Sub

Sub aaGeneralEXEC()
    Call progressIndicator_Begin("Executive Summary")
    pb.AddProgress 10
    If Range("executive_summary").Value = "Yes" Then
        pb.AddCaption "Creating Executive Summary..."
        Sheets("execSum").Visible = True
        Call execparts
        pb.AddProgress 65
        Call execpage
        pb.AddProgress 25
    Else
        If Sheets("execSum").Visible = True Then
            Sheets("execSum").Visible = False
        End If
    End If
    
    Call progressIndicator_End
    Worksheets("dashboard").Activate
End Sub

