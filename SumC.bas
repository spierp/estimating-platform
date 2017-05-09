Attribute VB_Name = "SumC"
Sub execparts()

'main box

Worksheets("execParts").Activate
Sheets("execParts").Cells.Select
Selection.EntireColumn.Hidden = False
Selection.EntireRow.Hidden = False
Cells(1, 1).Select

   Dim totalref As String
   If range("trade_detail").Value = "Yes" And _
    Sheets("Data").range("dataTable[[#Totals],[CONTRACT ITEM]]").Value > 0 Then
        totalref = "trade"
    ElseIf range("uniformat_item_detail").Value = "Yes" And _
    Sheets("Data").range("dataTable[[#Totals],[UNI L2]]").Value > 0 Then
        totalref = "uni2"
    ElseIf range("uniformat_item_detail").Value = "Yes" And _
    Sheets("Data").range("dataTable[[#Totals],[UNI L3/L4]]").Value > 0 Then
        totalref = "uni34"
    Else: MsgBox ("you need to provide contract item or Uniformat tags in order to generate a total project cost.")
    End If
    
Dim zonenumber As Integer
Dim zone As Integer
Dim y As Integer
zone = 1
zonenumber = WorksheetFunction.CountA(Sheets("dashboard").range("F22:Q22"))
'MsgBox (zonenumber)

Cells(8, 2).Value = range(totalref & "_total_cost").Value


If range("prim_div_qty").Value > 0 Then
    Cells(9, 2).Value = range(totalref & "_total_cost").Value / range("prim_div_qty").Value
Else
    Rows(3).EntireRow.Hidden = True
    Rows(9).EntireRow.Hidden = True
End If

If range("sec_div_qty").Value > 0 Then
    Cells(10, 2).Value = range(totalref & "_total_cost").Value / range("sec_div_qty").Value
Else
    Rows(4).EntireRow.Hidden = True
    Rows(10).EntireRow.Hidden = True
End If

If range("count_qty").Value > 0 Then
    Cells(11, 2).Value = range(totalref & "_total_cost").Value / range("count_qty").Value
Else
    Rows(5).EntireRow.Hidden = True
    Rows(11).EntireRow.Hidden = True
End If

If range("dur_qty").Value > 0 Then
    Cells(12, 2).Value = range(totalref & "_total_cost").Value / range("dur_qty").Value
Else
    Rows(6).EntireRow.Hidden = True
    Rows(12).EntireRow.Hidden = True
End If


For y = 3 To zonenumber + 2 Step 1
    Cells(8, y).Value = range(totalref & "_total_cost_Z" & zone)
    If range("prim_div_qty_Z1").Value > 0 Then
        Cells(9, y).Value = range(totalref & "_total_cost_Z" & zone) / range("prim_div_qty_Z" & zone).Value
    End If
    If range("sec_div_qty").Value > 0 Then
        Cells(10, y).Value = range(totalref & "_total_cost_Z" & zone) / range("sec_div_qty_Z" & zone).Value
    End If
    If range("count_Z1").Value > 0 Then
        Cells(11, y).Value = range(totalref & "_total_cost_Z" & zone) / range("count_Z" & zone).Value
    End If
    If range("dur_Z1").Value > 0 Then
        Cells(12, y).Value = range(totalref & "_total_cost_Z" & zone) / range("dur_Z" & zone).Value
    End If
    zone = zone + 1
Next y

For y = 1 To 14 Step 1
    If Cells(1, y).Value = "0" Then
        Columns(y).EntireColumn.Hidden = True
    End If
Next y
    
'trade graph

Dim btm As String

Sheets("tradeSum").Activate

btm = Sheets("tradeSum").Columns(3).Find("COST OF WORK - SUBTOTAL").Offset(-1, 1).Address

range("C12:" & btm).Copy

Sheets("execParts").Activate

Sheets("execParts").range("AB52").Select
ActiveSheet.Paste Link:=True

Application.CutCopyMode = False
ActiveWorkbook.Worksheets("execParts").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("execParts").Sort.SortFields.Add Key:=range( _
        "AC52:AC150"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("execParts").Sort
        .SetRange Selection
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Dim bottom As Integer
Dim toptradesrng As String
Dim x As Integer

For x = 150 To 52 Step -1
    If Cells(x, 29).Value <> "Excl." And Cells(x, 29).Value > 0 Then
        bottom = x
        x = 52
    End If
Next x

toptradesrng = "execParts!$AB$" & bottom - 9 & ":" & "$AC$" & bottom

ActiveSheet.ChartObjects("TopTrades").Activate
ActiveChart.SetSourceData Source:=range(toptradesrng)

'zonePie

Dim zonepierng As String

zonepierng = "execParts!$C$1:" & Cells(1, zonenumber + 2).Address & ",execParts!$C$8:" & Cells(8, zonenumber + 2).Address
ActiveSheet.ChartObjects("ZonePie").Activate
ActiveChart.SetSourceData Source:=range(zonepierng)

'zonePie

Dim zoneprimdivrng As String

zoneprimdivrng = "execParts!$C$1:" & Cells(1, zonenumber + 2).Address & ",execParts!$C$9:" & Cells(9, zonenumber + 2).Address
ActiveSheet.ChartObjects("ZonePrimDiv").Activate
ActiveChart.SetSourceData Source:=range(zoneprimdivrng)

End Sub

Sub execpage()
Sheets("execSum").Activate
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.15)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        If ThisWorkbook.Names("page_orientation").RefersToRange(1, 1).Value = "Portrait" Then
            .Orientation = xlPortrait
        Else
            .Orientation = xlLandscape
        End If
        .Draft = False
        If ThisWorkbook.Names("page_size").RefersToRange(1, 1).Value = "Letter" Then
            .PaperSize = xlPaperLetter
        ElseIf ThisWorkbook.Names("page_size").RefersToRange(1, 1).Value = "Legal" Then
            .PaperSize = xlPaperLegal
        Else
            .PaperSize = xlPaperTabloid
        End If
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
    End With
    Application.PrintCommunication = True

End Sub
