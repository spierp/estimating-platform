Attribute VB_Name = "AMasterSum"
Sub aaSumSheets()
'    Call OptimizeCode_Begin

'SUMMARY TAB
    If Range("trade_summary").Value = "Yes" And _
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
    
    If Range("uniformat_L2_summary").Value = "Yes" And _
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
    
    If Range("uniformat_L34_summary").Value = "Yes" And _
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
    
    If Range("breakouts_summary").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(3).Total > 1 Then
        Sheets("brkSum").Visible = True
        Call progressIndicator_Begin("Break-Outs Summary")
        Call sumData("brkSum")
        Call sumColumns
        Call sumPageSetup
        Call progressIndicator_End
    Else
        If Sheets("brkSum").Visible = True Then
            Sheets("brkSum").Visible = False
        End If
    End If
    
    If Range("alternates_summary").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(4).Total > 1 Then
        Sheets("altSum").Visible = True
        Call progressIndicator_Begin("Alternates Summary")
        Call sumData("altSum")
        Call sumColumns
        Call sumPageSetup
        Call progressIndicator_End
    Else
        If Sheets("altSum").Visible = True Then
            Sheets("altSum").Visible = False
        End If
    End If
       
    Worksheets("dashboard").Activate
    
'    Call sumErrorCheck
'    Call OptimizeCode_End

End Sub
