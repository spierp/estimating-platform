Attribute VB_Name = "AMasterSum"
Sub aaSumSheets()
'    Call OptimizeCode_Begin

'summary tabs

    If range("trade_summary").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(10).Total > 3 Then
        Sheets("tradeSum").Visible = True
        Call progressIndicator_Begin("Trade Summary Report")
        Call sumData("tradeSum")
        Call sumColumns
        Call sumPageSetup
        Call progressIndicator_End
    Else
        If Sheets("tradeSum").Visible = True Then
            Sheets("TradeSum").Visible = False
        End If
    End If
    
    If range("uniformat_L2_summary").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(8).Total > 3 Then
        Sheets("uni2Sum").Visible = True
        Call progressIndicator_Begin("System Summary Report")
        Call sumData("uni2Sum")
        Call sumColumns
        Call sumPageSetup
        Call progressIndicator_End
    Else
        If Sheets("uni2Sum").Visible = True Then
            Sheets("uni2Sum").Visible = False
        End If
    End If
    If range("uniformat_L34_summary").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(9).Total > 3 Then
        Sheets("uni34Sum").Visible = True
        Call progressIndicator_Begin("UniFormat Level 4 Summary Report")
        Call sumData("uni34Sum")
        Call sumColumns
        Call sumPageSetup
        Call progressIndicator_End
    Else
        If Sheets("uni34Sum").Visible = True Then
            Sheets("uni34Sum").Visible = False
        End If
    End If
    
    If range("trade_variance").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(10).Total > 3 Then
        Sheets("tradeVar").Visible = True
        Call progressIndicator_Begin("Trade Variance Report")
        Call variance("tradeVar")
        Call sumPageSetup
        Call progressIndicator_End
    Else
        If Sheets("tradeVar").Visible = True Then
            Sheets("TradeVar").Visible = False
        End If
    End If
    
'general tabs
    
    Call progressIndicator_Begin("General Tabs")
    pb.AddCaption "Working on general tabs..."
    If range("coverpage").Value = "Yes" Then
        pb.AddCaption "Formatting coverpage..."
        Sheets("cover").Visible = True
        Call coverPage
    Else
        If Sheets("cover").Visible = True Then
            Sheets("cover").Visible = False
        End If
    End If
    pb.AddProgress 10
    
    If range("tablecontents").Value = "Yes" Then
        pb.AddCaption "Creating Table of Contents..."
        Sheets("TOC").Visible = True
        Call tableofContents
    Else
        If Sheets("TOC").Visible = True Then
            Sheets("TOC").Visible = False
        End If
    End If
    pb.AddProgress 20
    
    If range("notesquals").Value = "Yes" Then
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
    
    If range("bim").Value > "0" Then
        pb.AddCaption "Creating BIM tabs..."
        Call BIM
    Else
        If Sheets("BIM-1").Visible = True Then
            Sheets("BIM-1").Visible = False
        End If
    End If
    pb.AddProgress 20
    
    If range("executive_summary").Value = "Yes" Then
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
    '    Call sumErrorCheck
'    Call OptimizeCode_End
End Sub
