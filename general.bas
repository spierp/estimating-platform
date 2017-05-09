Attribute VB_Name = "general"
Sub tableofContents()
    Sheets("TOC").Activate
    Sheets("TOC").Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
    Cells(1, 1).Select

Dim pageCount As Integer
pageCount = 1

'coverstuff
If ThisWorkbook.Names("coverpage").RefersToRange(1, 1).Value = "Yes" Then
    Rows("8").EntireRow.Hidden = False
    pageCount = pageCount + 1
Else
    Rows("8").EntireRow.Hidden = True
End If

If ThisWorkbook.Names("tablecontents").RefersToRange(1, 1).Value = "Yes" Then
    Rows("9").EntireRow.Hidden = False
    pageCount = pageCount + 1
Else
    Rows("9").EntireRow.Hidden = True
End If

'summaries
If range("trade_summary").Value = "Yes" Or range("executive_summary").Value = "Yes" _
Or range("uniformat_L2_summary").Value = Yes Or range("uniformat_L34_summary").Value = Yes Then
    Rows("10").EntireRow.Hidden = False
Else: Rows("10").EntireRow.Hidden = True
End If

If range("executive_summary").Value = "Yes" Then
    Rows("11").EntireRow.Hidden = False
    range("E11").Value = pageCount
    pageCount = pageCount + Sheets("execSum").PageSetup.Pages.count
Else: Rows("11").EntireRow.Hidden = True
End If

If range("trade_summary").Value = "Yes" Then
    Rows("12").EntireRow.Hidden = False
    range("E12").Value = pageCount
    pageCount = pageCount + Sheets("tradeSum").PageSetup.Pages.count
Else: Rows("12").EntireRow.Hidden = True
End If

If range("uniformat_L2_summary").Value = "Yes" Then
    Rows("13").EntireRow.Hidden = False
    range("E13").Value = pageCount
    pageCount = pageCount + Sheets("uni2Sum").PageSetup.Pages.count
Else: Rows("13").EntireRow.Hidden = True
End If

If range("uniformat_L34_summary").Value = "Yes" Then
    Rows("14").EntireRow.Hidden = False
    range("E14").Value = pageCount
    pageCount = pageCount + Sheets("uni34Sum").PageSetup.Pages.count
Else: Rows("14").EntireRow.Hidden = True
End If

'notesquals
If range("notesquals").Value = "Yes" Then
    Rows("15:16").EntireRow.Hidden = False
    range("E16").Value = pageCount
    pageCount = pageCount + Sheets("N+Q").PageSetup.Pages.count
Else: Rows("15:16").EntireRow.Hidden = True
End If

'variance
If range("trade_variance").Value = "Yes" Or range("uniformat_L2_variance").Value = "Yes" Or range("uniformat_L34_variance").Value = "Yes" Then
    Rows("17").EntireRow.Hidden = False
Else: Rows("17").EntireRow.Hidden = True
End If

If range("trade_variance").Value = "Yes" Then
    Rows("18").EntireRow.Hidden = False
    range("E18").Value = pageCount
    pageCount = pageCount + 1 'Sheets("tradeVar").PageSetup.Pages.Count
Else: Rows("18").EntireRow.Hidden = True
End If

If range("uniformat_L2_variance").Value = "Yes" Then
    Rows("19").EntireRow.Hidden = False
    range("E19").Value = pageCount
    pageCount = pageCount + 1 'Sheets("uni2Var").PageSetup.Pages.Count
Else: Rows("19").EntireRow.Hidden = True
End If

If range("uniformat_L34_variance").Value = "Yes" Then
    Rows("20").EntireRow.Hidden = False
    range("E20").Value = pageCount
    pageCount = pageCount + 1 'Sheets("uni34Var").PageSetup.Pages.Count
Else: Rows("20").EntireRow.Hidden = True
End If

'breakouts and alts
If range("breakouts_detail").Value = "Yes" Or range("alternates_detail").Value = "Yes" Then
    Rows("21").EntireRow.Hidden = False
Else: Rows("21").EntireRow.Hidden = True
End If

If range("breakouts_detail").Value = "Yes" Then
    Rows("22:23").EntireRow.Hidden = False
    range("E22").Value = pageCount
    pageCount = pageCount + Sheets("brkSum").PageSetup.Pages.count
    range("E23").Value = pageCount
    pageCount = pageCount + Sheets("brkDetail").PageSetup.Pages.count
Else: Rows("22:23").EntireRow.Hidden = True
End If

If range("alternates_detail").Value = "Yes" Then
    Rows("24:25").EntireRow.Hidden = False
    range("E24").Value = pageCount
    pageCount = pageCount + 1 'Sheets("altSum").PageSetup.Pages.Count
    range("E25").Value = pageCount
    pageCount = pageCount + Sheets("brkDetail").PageSetup.Pages.count
Else: Rows("24:25").EntireRow.Hidden = True
End If

If range("bim").Value > "0" Then
    Rows("26:27").EntireRow.Hidden = False
    range("E27").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("26:27").EntireRow.Hidden = True
End If

If range("bim").Value > "1" Then
    Rows("28").EntireRow.Hidden = False
    range("E28").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("28").EntireRow.Hidden = True
End If

If range("bim").Value > "2" Then
    Rows("29").EntireRow.Hidden = False
    range("E29").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("29").EntireRow.Hidden = True
End If

If range("bim").Value > "3" Then
    Rows("30").EntireRow.Hidden = False
    range("E30").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("30").EntireRow.Hidden = True
End If

If range("bim").Value > "4" Then
    Rows("31").EntireRow.Hidden = False
    range("E31").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("31").EntireRow.Hidden = True
End If

If range("bim").Value > "5" Then
    Rows("32").EntireRow.Hidden = False
    range("E32").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("32").EntireRow.Hidden = True
End If

If range("bim").Value > "6" Then
    Rows("33").EntireRow.Hidden = False
    range("E33").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("33").EntireRow.Hidden = True
End If

If range("bim").Value > "7" Then
    Rows("34").EntireRow.Hidden = False
    range("E34").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("34").EntireRow.Hidden = True
End If

If range("bim").Value > "7" Then
    Rows("34").EntireRow.Hidden = False
    range("E34").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("34").EntireRow.Hidden = True
End If

If range("trade_detail").Value = "Yes" Or range("uniformat_item_detail").Value = "Yes" Then
    Rows("35").EntireRow.Hidden = False
Else: Rows("35").EntireRow.Hidden = True
End If

If range("trade_detail").Value = "Yes" Then
    Rows("36").EntireRow.Hidden = False
    range("E36").Value = pageCount
    pageCount = pageCount + Sheets("tradeDetail").PageSetup.Pages.count
Else: Rows("36").EntireRow.Hidden = True
End If

If range("uniformat_item_detail").Value = "Yes" Then
    Rows("37").EntireRow.Hidden = False
    range("E37").Value = pageCount
    pageCount = pageCount + Sheets("uniDetail").PageSetup.Pages.count
Else: Rows("37").EntireRow.Hidden = True
End If

End Sub

Sub coverPage()
Worksheets("cover").Activate

If range("page_orientation").Value = "Portrait" Then
    Columns("B").ColumnWidth = 15
    Columns("C").ColumnWidth = 100
    Rows("36").RowHeight = 160
    Rows("44").RowHeight = 130
ElseIf range("page_orientation").Value = "Landscape" Then
    Columns("B").ColumnWidth = 45
    Columns("C").ColumnWidth = 108
    Rows("36").RowHeight = 12.75
    Rows("44").RowHeight = 50
End If

If range("page_size").Value = "Letter" And range("page_orientation").Value = "Landscape" Then
    Columns("B").ColumnWidth = 45
ElseIf range("page_size").Value = "Tabloid" And range("page_orientation").Value = "Landscape" Then
    Columns("B").ColumnWidth = 85
End If


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

Sub notesQualsCopy()
Attribute notesQualsCopy.VB_ProcData.VB_Invoke_Func = " \n14"
Sheets("clipboard").Visible = True
Sheets("nqParts").Visible = True

Worksheets("nqParts").Cells.ClearContents
Worksheets("nqParts").Cells.Borders.LineStyle = xlNone
If range("page_size").Value = "Tabloid" And range("page_orientation").Value = "Landscape" Then
    Worksheets("nqParts").Columns("E").ColumnWidth = 120
Else
    Worksheets("nqParts").Columns("E").ColumnWidth = 84
End If

Worksheets("N+Q-Data").range("A1").CurrentRegion.Copy
Worksheets("clipboard").range("M1").PasteSpecial _
    Paste:=xlPasteValues, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=True

Application.CutCopyMode = False


Worksheets("clipboard").Activate
Dim nqrow As Integer
nqrow = -1

Dim yy As Integer
Dim zz As Integer
Dim rccc As Integer

For yy = 13 To 56 Step 1
    If Cells(1, yy).Value <> "" Then
        
        'copy main category
        nqrow = nqrow + 2
        Cells(1, yy).Copy
        Worksheets("nqParts").Cells(nqrow, 2).PasteSpecial _
            Paste:=xlPasteValues, Operation:= _
            xlNone, SkipBlanks:=True, Transpose:=False
        With Worksheets("nqParts").range(Cells(nqrow, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & _
        Cells(nqrow, 5).Address(RowAbsolute:=False, ColumnAbsolute:=False)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        nqrow = nqrow + 2

        
        For zz = yy To 56 Step 1
            If Cells(1, zz) = "" And Cells(3, zz) <> "" Then
                If Cells(2, zz - 1) = "" And zz - 1 > yy Then
                    zz = 56
                Else:
                    rcc = Cells(Rows.count, zz).End(xlUp).Row - 2
'                    nqrow = nqrow + 1
                    Cells(2, zz).Copy
                    Worksheets("nqParts").Activate
                    Cells(nqrow, 3).PasteSpecial _
                        Paste:=xlPasteValues, Operation:= _
                        xlNone, SkipBlanks:=True, Transpose:=False
                        nqrow = nqrow + 1
                    Worksheets("clipboard").Activate
                    range(Cells(3, zz).Address & ":" & _
                    Cells(2 + rcc, zz).Address).Copy
                    Worksheets("nqParts").Activate
                    Cells(nqrow, 5).PasteSpecial _
                        Paste:=xlPasteValues, Operation:= _
                        xlNone, SkipBlanks:=True, Transpose:=False
                        nqrow = nqrow + rcc
                    Worksheets("clipboard").Activate
                End If
            End If
        Next zz
        
    End If
Next yy

Dim dat As Variant
Dim rng As range
Dim i As Long

Worksheets("nqParts").Activate
Set rng = range("$D$1", Cells(Rows.count, "E").End(xlUp)).Cells
dat = rng.Value

For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 2) <> "" Then
        dat(i, 1) = "•"
    End If
Next
rng.Value = dat
Application.CutCopyMode = False
'Sheets("clipboard").Visible = False

End Sub

Sub notesQualsInsert()

'delete existing notes
Dim pic As Variant
For Each pic In Worksheets("N+Q").Pictures
    If pic.Name = "nqPic1" Or pic.Name = "nqPic2" Then
        pic.Delete
    End If
Next

Worksheets("nqParts").Activate

'find first row of right column
Dim firstnr As Integer
Dim lastrw As Integer
Dim x As Integer
Dim btm As Integer
lastrw = Worksheets("nqParts").Cells(Rows.count, "E").End(xlUp).Row
firstnr = 0

If range("page_orientation").Value = "Landscape" And range("page_size").Value = "Letter" Then
    btm = 700
ElseIf range("page_orientation").Value = "Landscape" And range("page_size").Value = "Tabloid" Then
    btm = 890
ElseIf range("page_orientation").Value = "Portrait" And range("page_size").Value = "Letter" Then
    btm = 700
End If

For x = 1 To lastrw Step 1
    If Worksheets("nqParts").range("A" & x).top > btm Then
        firstnr = (Cells(x, 1).Row)
        'MsgBox (firstnr)
        x = lastrw
    End If
Next x

If firstnr <> "0" Then
    If Cells(firstnr - 1, 3) <> "" Or Cells(firstnr - 1, 2) <> "" Then
        firstnr = firstnr - 1
    ElseIf Cells(firstnr - 2, 2) <> "" Then
        firstnr = firstnr - 2
    ElseIf Cells(firstnr - 3, 2) <> "" Then
        firstnr = firstnr - 3
    End If
End If

'find header of right column
'Dim rng1 As range
'Set rng1 = Worksheets("nqParts").range("A1:F" & firstnr - 1).Find("*", Worksheets("nqParts").[b1], xlValues, , xlByRows, xlPrevious)
'If Not rng1 Is Nothing Then
'    MsgBox "last cell is " & rng1.Address(0, 0)
'Else
'    MsgBox Worksheets("nqParts").Name & " columns B:B are empty", vbCritical
'End If

'MsgBox (firstnr)

'copy data to N+Q tab
Dim nqPic1 As Picture

If firstnr = "0" Then
    Worksheets("nqParts").range("A1:F" & lastrw).Copy
Else
    Worksheets("nqParts").range("A1:F" & firstnr - 1).Copy
End If

With Worksheets("N+Q")
    Set nqPic1 = .Pictures.Paste(Link:=True)
    nqPic1.Name = "nqPic1"
    nqPic1.Left = .range("A8").Left
    nqPic1.top = .range("A8").top
End With

Dim nqPic2 As Picture

If firstnr <> "0" Then
Worksheets("nqParts").range("A" & firstnr & ":F" & lastrw).Copy

With Worksheets("N+Q")
    Set nqPic1 = .Pictures.Paste(Link:=True)
    nqPic1.Name = "nqPic2"
    nqPic1.Left = .range("E8").Left
    nqPic1.top = .range("E8").top
End With
End If

Application.CutCopyMode = False
Sheets("nqParts").Visible = False
End Sub

Sub notesQualsFormat()

Sheets("N+Q").Activate
'If range("page_orientation").Value = "Portrait" Then
'    Columns("B").ColumnWidth = 15
'    Columns("C").ColumnWidth = 100
'    Rows("36").RowHeight = 160
'    Rows("44").RowHeight = 130
'ElseIf range("page_orientation").Value = "Landscape" Then
'    Columns("B").ColumnWidth = 45
'    Columns("C").ColumnWidth = 108
'    Rows("36").RowHeight = 12.75
'    Rows("44").RowHeight = 50
'End If

If range("page_size").Value = "Letter" And range("page_orientation").Value = "Landscape" Then
    Sheets("N+Q").Columns("C").ColumnWidth = 85
    Sheets("N+Q").Columns("E").ColumnWidth = 85
    Sheets("N+Q").Rows("8:10").RowHeight = 235
ElseIf range("page_size").Value = "Tabloid" And range("page_orientation").Value = "Landscape" Then
    Sheets("N+Q").Columns("C").ColumnWidth = 130
    Sheets("N+Q").Columns("E").ColumnWidth = 130
    Sheets("N+Q").Rows("8:10").RowHeight = 300
End If

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

Sub BIM()
Dim bimArray As Variant
Dim i As Long

bimArray = Array("BIM-1", "BIM-2", "BIM-3", "BIM-4", "BIM-5", "BIM-6", "BIM-7", "BIM-8")

For i = 0 To UBound(bimArray)
    If i < range("bim").Value Then
        Sheets(bimArray(i)).Visible = True
        
        If range("page_size").Value = "Letter" And range("page_orientation").Value = "Landscape" Then

        ElseIf range("page_size").Value = "Tabloid" And range("page_orientation").Value = "Landscape" Then
            Sheets(bimArray(i)).Columns("C:D").ColumnWidth = 127
            Sheets(bimArray(i)).Rows("8:10").RowHeight = 295
        End If
        
        Application.PrintCommunication = False
        With Sheets(bimArray(i)).PageSetup
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
    Else
        If Sheets(bimArray(i)).Visible = True Then
            Sheets(bimArray(i)).Visible = False
        End If
    End If
    
Next i

End Sub
