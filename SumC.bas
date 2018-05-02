Attribute VB_Name = "SumC"
Sub execparts()

'MAIN BOX

Worksheets("execParts").Activate
Sheets("execParts").Cells.Select
Selection.EntireColumn.Hidden = False
Selection.EntireRow.Hidden = False
Cells(1, 1).Select

   Dim totalref As String
   If Range("trade_summary").Value = "Yes" And _
    Sheets("Data").Range("dataTable[[#Totals],[CONTRACT ITEM]]").Value > 0 Then
        totalref = "trade"
    ElseIf Range("uniformat_L2_summary").Value = "Yes" And _
    Sheets("Data").Range("dataTable[[#Totals],[UNI L2]]").Value > 0 Then
        totalref = "uni2"
'    ElseIf Range("uniformat_item_detail").Value = "Yes" And _
'    Sheets("Data").Range("dataTable[[#Totals],[UNI L3/L4]]").Value > 0 Then
'        totalref = "uni34"
    Else: MsgBox ("Error: #1 You need to provide contract item or Uniformat tags in order to generate a total project cost. #2 Make sure either Trade Summary or Uniformat L2 Summary report is enabled.  This is required to calculate mark-ups.")
    End If
    
Dim zonenumber As Integer
Dim zone As Integer
Dim y As Integer
zone = 1
zonenumber = WorksheetFunction.CountA(Sheets("dashboard").Range("F23:Q23"))

'main box totals
    Cells(8, 2).Formula = "=" & totalref & "_total_cost"
    
    If Range("prim_div_qty").Value > 0 Then
        Cells(9, 2).Formula = "=" & totalref & "_total_cost" & "/prim_div_qty"
    Else
        Rows(3).EntireRow.Hidden = True
        Rows(9).EntireRow.Hidden = True
    End If
    
    If Range("sec_div_qty").Value > 0 Then
        Cells(10, 2).Formula = "=" & totalref & "_total_cost" & "/sec_div_qty"
    Else
        Rows(4).EntireRow.Hidden = True
        Rows(10).EntireRow.Hidden = True
    End If
    
    If Range("count_qty").Value > 0 Then
        Cells(11, 2).Formula = "=" & totalref & "_total_cost" & "/count_qty"
    Else
        Rows(5).EntireRow.Hidden = True
        Rows(11).EntireRow.Hidden = True
    End If
    
    If Range("dur_qty").Value > 0 Then
        Cells(12, 2).Formula = "=" & totalref & "_total_cost" & "/dur_qty"
    Else
        Rows(6).EntireRow.Hidden = True
        Rows(12).EntireRow.Hidden = True
    End If

'mainbox zones
For y = 3 To zonenumber + 2 Step 1
    Cells(8, y).Formula = "=" & totalref & "_total_cost_Z" & zone
    If Range("prim_div_qty_Z1").Value > 0 Then
        Cells(9, y).Formula = "=" & totalref & "_total_cost_Z" & zone & "/prim_div_qty_Z" & zone
    End If
    If Range("sec_div_qty").Value > 0 Then
        Cells(10, y).Formula = "=" & totalref & "_total_cost_Z" & zone & "/sec_div_qty_Z" & zone
    End If
    If Range("count_Z1").Value > 0 Then
        Cells(11, y).Formula = "=" & totalref & "_total_cost_Z" & zone & "/count_Z" & zone
    End If
    If Range("dur_Z1").Value > 0 Then
        Cells(12, y).Formula = "=" & totalref & "_total_cost_Z" & zone & "/dur_Z" & zone
    End If
    zone = zone + 1
Next y

For y = 1 To 14 Step 1
    If Cells(1, y).Value = "0" Then
        Columns(y).EntireColumn.Hidden = True
    End If
Next y
    
'BREAK-OUTS
Rows("13:51").EntireRow.Hidden = False
Dim breakcount As Integer
breakcount = Sheets("Data").Range("dataTable[[#Totals],[BRK]]").Value

If breakcount = 11 Then
    Rows("48:50").EntireRow.Hidden = True
ElseIf breakcount = 10 Then
    Rows("45:50").EntireRow.Hidden = True
ElseIf breakcount = 9 Then
    Rows("42:50").EntireRow.Hidden = True
ElseIf breakcount = 8 Then
    Rows("39:50").EntireRow.Hidden = True
ElseIf breakcount = 7 Then
    Rows("36:50").EntireRow.Hidden = True
ElseIf breakcount = 6 Then
    Rows("33:50").EntireRow.Hidden = True
ElseIf breakcount = 5 Then
    Rows("30:50").EntireRow.Hidden = True
ElseIf breakcount = 4 Then
    Rows("27:50").EntireRow.Hidden = True
ElseIf breakcount = 3 Then
    Rows("24:50").EntireRow.Hidden = True
ElseIf breakcount = 2 Then
    Rows("21:50").EntireRow.Hidden = True
ElseIf breakcount = 1 Then
    Rows("18:50").EntireRow.Hidden = True
ElseIf breakcount = 0 Then
    Rows("13:51").EntireRow.Hidden = True
End If

    
'TRADE GRAPH
Dim btm As String

Columns(28).ClearContents
Columns(29).ClearContents

Sheets("tradeSum").Activate

btm = Sheets("tradeSum").Columns(3).Find("COST OF WORK - SUBTOTAL").Offset(-1, 1).Address

Range("C12:" & btm).Copy

Sheets("execParts").Activate

Sheets("execParts").Range("AB78").Select
ActiveSheet.Paste Link:=True

Application.CutCopyMode = False
ActiveWorkbook.Worksheets("execParts").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("execParts").Sort.SortFields.Add Key:=Range( _
        "AC78:AC178"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
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

For x = 178 To 78 Step -1
    If Cells(x, 29).Value <> "Excl." And Cells(x, 29).Value > 0 Then
        bottom = x
        x = 78
    End If
Next x

toptradesrng = "execParts!$AB$" & bottom - 9 & ":" & "$AC$" & bottom

ActiveSheet.ChartObjects("TopTrades").Activate
ActiveChart.SetSourceData Source:=Range(toptradesrng)

'zonePie
Dim zonepierng As String

zonepierng = "execParts!$C$1:" & Cells(1, zonenumber + 2).Address & ",execParts!$C$8:" & Cells(8, zonenumber + 2).Address
ActiveSheet.ChartObjects("ZonePie").Activate
ActiveChart.SetSourceData Source:=Range(zonepierng)

'zoneDiv
Dim zoneprimdivrng As String

zoneprimdivrng = "execParts!$C$1:" & Cells(1, zonenumber + 2).Address & ",execParts!$C$9:" & Cells(9, zonenumber + 2).Address
ActiveSheet.ChartObjects("ZonePrimDiv").Activate
ActiveChart.SetSourceData Source:=Range(zoneprimdivrng)

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
        .FooterMargin = Application.InchesToPoints(0.17)
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
