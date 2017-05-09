Attribute VB_Name = "AMasterDetail"
Sub aaDetailSheets()

    Call purgeDetail
    
    If range("trade_detail").Value = "Yes" Then
        If Sheets("Data").range("dataTable[[#Totals],[CONTRACT ITEM]]").Value > 0 Then
            Call progressIndicator_Begin("Contract Item Detail Report")
            Call DetailTransfer("tradeDetail")
            Call sortAndFormat
            Call createSubTotals
            Call createHeadings
            Call tableFormat
            Call pageFormat
            Call progressIndicator_End
        Else
            MsgBox ("No Contract Item data found")
        End If
    End If
    
    If range("uniformat_item_detail").Value = "Yes" Then
        If Sheets("Data").range("dataTable[[#Totals],[UNI L2]]").Value > 0 Then
            Call progressIndicator_Begin("UniFormat Detail Report")
            Call DetailTransfer("uniDetail")
            Call sortAndFormat
            Call createSubTotals
            Call createHeadings
            Call tableFormat
            Call pageFormat
            Call progressIndicator_End
        Else
            MsgBox ("No UniFormat data found")
        End If
    End If
    
    If range("breakouts_detail").Value = "Yes" Then
        If Sheets("Data").range("dataTable[[#Totals],[BRK]]").Value > 0 Then
            Call progressIndicator_Begin("Break-Outs Detail Report")
            Call DetailTransfer("brkDetail")
            Call sortAndFormat
            Call createSubTotals
            Call createHeadings
            Call tableFormat
            Call pageFormat
            Call progressIndicator_End
        Else
            MsgBox ("No Break-Outs Found")
        End If
    End If

    If range("alternates_detail").Value = "Yes" Then
        If Sheets("Data").range("dataTable[[#Totals],[ALT]]").Value > 0 Then
            Call progressIndicator_Begin("Alternates Detail Report")
            Call DetailTransfer("altDetail")
            Call sortAndFormat
            Call createSubTotals
            Call createHeadings
            Call tableFormat
            Call pageFormat
            Call progressIndicator_End
        Else
            MsgBox ("No Alternates Found")
        End If
    End If
    
    Call detailTabs
    
    Worksheets("dashboard").Activate
    
End Sub



Sub purgeDetail()

Dim ws As Worksheet
For Each ws In Worksheets
    If ws.Name = "uniDetail" Then
        Application.DisplayAlerts = False
        Sheets("uniDetail").Delete
        Application.DisplayAlerts = True
    End If
Next
For Each ws In Worksheets
    If ws.Name = "tradeDetail" Then
        Application.DisplayAlerts = False
        Sheets("tradeDetail").Delete
        Application.DisplayAlerts = True
    End If
Next
For Each ws In Worksheets
    If ws.Name = "brkDetail" Then
        Application.DisplayAlerts = False
        Sheets("brkDetail").Delete
        Application.DisplayAlerts = True
    End If
Next
For Each ws In Worksheets
    If ws.Name = "altDetail" Then
        Application.DisplayAlerts = False
        Sheets("altDetail").Delete
        Application.DisplayAlerts = True
    End If

Next

End Sub

Sub detailTabs()
    
'    If range("alternates_detail").Value = "Yes" _
'    And Sheets("Data").range("dataTable[[#Totals],[ALT]]").Value > 0 Then
'        Sheets("altDetail").Select
'        Sheets("altDetail").Move After:=Sheets("altSum")
'        With ActiveWorkbook.Sheets("altDetail").Tab
'            .ThemeColor = xlThemeColorAccent4
'            .TintAndShade = 0.599993896298105
'        End With
'    End If
    If range("breakouts_detail").Value = "Yes" _
    And Sheets("Data").range("dataTable[[#Totals],[BRK]]").Value > 0 Then
        Sheets("brkDetail").Select
        Sheets("brkDetail").Move After:=Sheets("brkSum")
        With ActiveWorkbook.Sheets("brkDetail").Tab
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.599993896298105
        End With
    End If
    If range("trade_detail").Value = "Yes" And _
    Sheets("Data").range("dataTable[[#Totals],[CONTRACT ITEM]]").Value > 0 Then
        Sheets("tradeDetail").Select
        Sheets("tradeDetail").Move After:=Sheets("BIM-8")
        With ActiveWorkbook.Sheets("tradeDetail").Tab
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
        End With
    End If
    If range("uniformat_item_detail").Value = "Yes" And _
    Sheets("Data").range("dataTable[[#Totals],[UNI L2]]").Value > 0 Then
        Sheets("uniDetail").Select
        Sheets("uniDetail").Move After:=Sheets("BIM-8")
        With ActiveWorkbook.Sheets("uniDetail").Tab
            .ThemeColor = xlThemeColorAccent4
            .TintAndShade = 0.399975585192419
        End With
    End If

End Sub



