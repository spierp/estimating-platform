Attribute VB_Name = "general"
Sub tableofContents()
    Sheets("TOC").Activate
    Sheets("TOC").Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
    Worksheets("TOC").Range("E8:E40").ClearContents
    Cells(1, 1).Select

Dim ws As Worksheet

Dim brkdetailsheet As Boolean
Dim altdetailsheet As Boolean
Dim unidetailsheet As Boolean
Dim tradedetailsheet As Boolean
Dim vardetailsheet As Boolean

For Each ws In Worksheets
    If ws.Name = "brkDetail" Then
        brkdetailsheet = True
    ElseIf ws.Name = "altDetail" Then
        altdetailsheet = True
    ElseIf ws.Name = "tradeDetail" Then
        tradedetailsheet = True
    ElseIf ws.Name = "uniDetail" Then
        unidetailsheet = True
    ElseIf ws.Name = "varDetail" Then
        vardetailsheet = True
    End If
Next

Dim pageCount As Integer
pageCount = 1

'COVERSTUFF
If Range("coverpage").Value = "Yes" Then
    Rows("8").EntireRow.Hidden = False
    pageCount = pageCount + 1
Else
    Rows("8").EntireRow.Hidden = True
End If

If Range("tablecontents").Value = "Yes" Then
    Rows("9").EntireRow.Hidden = False
    pageCount = pageCount + 1
Else
    Rows("9").EntireRow.Hidden = True
End If

'SUMMARIES
If Range("trade_summary").Value = "Yes" Or Range("executive_summary").Value = "Yes" _
Or Range("uniformat_L2_summary").Value = yes Or Range("uniformat_L34_summary").Value = yes Then
    Rows("10").EntireRow.Hidden = False
Else: Rows("10").EntireRow.Hidden = True
End If

If Range("executive_summary").Value = "Yes" Then
    Rows("11").EntireRow.Hidden = False
    Range("E11").Value = pageCount
    pageCount = pageCount + Sheets("execSum").PageSetup.Pages.Count
Else: Rows("11").EntireRow.Hidden = True
End If

If Range("trade_summary").Value = "Yes" Then
    Rows("12").EntireRow.Hidden = False
    Range("E12").Value = pageCount
    pageCount = pageCount + Sheets("tradeSum").PageSetup.Pages.Count
Else: Rows("12").EntireRow.Hidden = True
End If

If Range("uniformat_L2_summary").Value = "Yes" Then
    Rows("13").EntireRow.Hidden = False
    Range("E13").Value = pageCount
    pageCount = pageCount + Sheets("uni2Sum").PageSetup.Pages.Count
Else: Rows("13").EntireRow.Hidden = True
End If

If Range("uniformat_L34_summary").Value = "Yes" Then
    Rows("14").EntireRow.Hidden = False
    Range("E14").Value = pageCount
    pageCount = pageCount + Sheets("uni34Sum").PageSetup.Pages.Count
Else: Rows("14").EntireRow.Hidden = True
End If

'NOTES & QUALS
If Range("notesquals").Value = "Yes" Then
    Rows("15:16").EntireRow.Hidden = False
    Range("E16").Value = pageCount
    pageCount = pageCount + Sheets("N+Q").PageSetup.Pages.Count
Else: Rows("15:16").EntireRow.Hidden = True
End If

'VARIANCE
If Range("trade_variance").Value = "Yes" Or Range("uniformat_L2_variance").Value = "Yes" Or Range("uniformat_L34_variance").Value = "Yes" Then
    Rows("17").EntireRow.Hidden = False
Else: Rows("17").EntireRow.Hidden = True
End If

If Range("trade_variance").Value = "Yes" Then
    Rows("18").EntireRow.Hidden = False
    Range("E18").Value = pageCount
    pageCount = pageCount + Sheets("tradeVar").PageSetup.Pages.Count
Else: Rows("18").EntireRow.Hidden = True
End If

If Range("uniformat_L2_variance").Value = "Yes" Then
    Rows("19").EntireRow.Hidden = False
    Range("E19").Value = pageCount
    pageCount = pageCount + Sheets("uni2Var").PageSetup.Pages.Count
Else: Rows("19").EntireRow.Hidden = True
End If

If Range("uniformat_L34_variance").Value = "Yes" Then
    Rows("20").EntireRow.Hidden = False
    Range("E20").Value = pageCount
    pageCount = pageCount + Sheets("uni34Var").PageSetup.Pages.Count
Else: Rows("20").EntireRow.Hidden = True
End If

'BREAK-OUTS & ALTERNATES
If Range("breakouts_detail").Value = "Yes" Or Range("alternates_detail").Value = "Yes" Or Range("breakouts_summary").Value = "Yes" Then
    Rows("21").EntireRow.Hidden = False
Else: Rows("21").EntireRow.Hidden = True
End If

'BREAK-OUTS

If Range("breakouts_summary").Value = "Yes" Then
    Rows("22").EntireRow.Hidden = False
    Range("E22").Value = pageCount
    pageCount = pageCount + Sheets("brkSum").PageSetup.Pages.Count
Else: Rows("22").EntireRow.Hidden = True
End If

If Range("breakouts_detail").Value = "Yes" And brkdetailsheet = True Then
    Rows("23").EntireRow.Hidden = False
    Range("E23").Value = pageCount
    pageCount = pageCount + Sheets("brkDetail").PageSetup.Pages.Count
Else: Rows("23").EntireRow.Hidden = True
End If

'ALTERNATES

If Range("alternates_summary").Value = "Yes" Then
    Rows("24").EntireRow.Hidden = False
    Range("E24").Value = pageCount
    pageCount = pageCount + Sheets("brkSum").PageSetup.Pages.Count
Else: Rows("24").EntireRow.Hidden = True
End If

If Range("alternates_detail").Value = "Yes" And altdetailsheet = True Then
    Rows("25").EntireRow.Hidden = False
    Range("E25").Value = pageCount
    pageCount = pageCount + Sheets("altDetail").PageSetup.Pages.Count
    Else: Rows("25").EntireRow.Hidden = True
End If

'BIM
If Range("bim").Value > "0" Then
    Rows("26:27").EntireRow.Hidden = False
    Range("E27").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("26:27").EntireRow.Hidden = True
End If

If Range("bim").Value > "1" Then
    Rows("28").EntireRow.Hidden = False
    Range("E28").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("28").EntireRow.Hidden = True
End If

If Range("bim").Value > "2" Then
    Rows("29").EntireRow.Hidden = False
    Range("E29").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("29").EntireRow.Hidden = True
End If

If Range("bim").Value > "3" Then
    Rows("30").EntireRow.Hidden = False
    Range("E30").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("30").EntireRow.Hidden = True
End If

If Range("bim").Value > "4" Then
    Rows("31").EntireRow.Hidden = False
    Range("E31").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("31").EntireRow.Hidden = True
End If

If Range("bim").Value > "5" Then
    Rows("32").EntireRow.Hidden = False
    Range("E32").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("32").EntireRow.Hidden = True
End If

If Range("bim").Value > "6" Then
    Rows("33").EntireRow.Hidden = False
    Range("E33").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("33").EntireRow.Hidden = True
End If

If Range("bim").Value > "7" Then
    Rows("34").EntireRow.Hidden = False
    Range("E34").Value = pageCount
    pageCount = pageCount + 1
Else: Rows("34").EntireRow.Hidden = True
End If

'DETAIL
If Range("trade_detail").Value = "Yes" Or Range("uniformat_item_detail").Value = "Yes" Or Range("detail_variance").Value = "Yes" Then
    Rows("35").EntireRow.Hidden = False
Else: Rows("35").EntireRow.Hidden = True
End If

If Range("trade_detail").Value = "Yes" And tradedetailsheet = True Then
    Rows("36").EntireRow.Hidden = False
    Range("E36").Value = pageCount
    pageCount = pageCount + Sheets("tradeDetail").PageSetup.Pages.Count
Else: Rows("36").EntireRow.Hidden = True
End If

If Range("uniformat_item_detail").Value = "Yes" And unidetailsheet = True Then
    Rows("37").EntireRow.Hidden = False
    Range("E37").Value = pageCount
    pageCount = pageCount + Sheets("uniDetail").PageSetup.Pages.Count
Else: Rows("37").EntireRow.Hidden = True
End If

If Range("detail_variance").Value = "Yes" And vardetailsheet = True Then
    Rows("38").EntireRow.Hidden = False
    Range("E38").Value = pageCount
    pageCount = pageCount + Sheets("varDetail").PageSetup.Pages.Count
Else: Rows("38").EntireRow.Hidden = True
End If


End Sub

Sub coverPage()
Worksheets("cover").Activate

If Range("page_orientation").Value = "Portrait" Then
    Columns("B").ColumnWidth = 15
    Columns("C").ColumnWidth = 100
    Rows("36").RowHeight = 160
    Rows("44").RowHeight = 130
ElseIf Range("page_orientation").Value = "Landscape" Then
    Columns("B").ColumnWidth = 45
    Columns("C").ColumnWidth = 108
    Rows("36").RowHeight = 12.75
    Rows("44").RowHeight = 50
End If

If Range("page_size").Value = "Letter" And Range("page_orientation").Value = "Landscape" Then
    Columns("B").ColumnWidth = 45
ElseIf Range("page_size").Value = "Tabloid" And Range("page_orientation").Value = "Landscape" Then
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

Sub notesQualsCopy()
Attribute notesQualsCopy.VB_ProcData.VB_Invoke_Func = " \n14"
Sheets("clipboard").Visible = True

Sheets("clipboard").Range("M1:BC25").Clear

Sheets("nqParts").Visible = True

Worksheets("nqParts").Cells.ClearContents
Worksheets("nqParts").Cells.Borders.LineStyle = xlNone

If Range("page_size").Value = "Tabloid" And Range("page_orientation").Value = "Landscape" Then
    Worksheets("nqParts").Columns("E").ColumnWidth = 120
Else
    Worksheets("nqParts").Columns("E").ColumnWidth = 84
End If

Worksheets("N+Q-Data").Range("A1").CurrentRegion.Copy
Worksheets("clipboard").Range("M1").PasteSpecial _
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
        With Worksheets("nqParts").Range(Cells(nqrow, 2).Address(RowAbsolute:=False, ColumnAbsolute:=False) & ":" & _
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
                    rcc = Cells(Rows.Count, zz).End(xlUp).Row - 2
'                    nqrow = nqrow + 1
                    Cells(2, zz).Copy
                    Worksheets("nqParts").Activate
                    Cells(nqrow, 3).PasteSpecial _
                        Paste:=xlPasteValues, Operation:= _
                        xlNone, SkipBlanks:=True, Transpose:=False
                        nqrow = nqrow + 1
                    Worksheets("clipboard").Activate
                    Range(Cells(3, zz).Address & ":" & _
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
Dim rng As Range
Dim i As Long

Worksheets("nqParts").Activate
Set rng = Range("$D$1", Cells(Rows.Count, "E").End(xlUp)).Cells
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
lastrw = Worksheets("nqParts").Cells(Rows.Count, "E").End(xlUp).Row
firstnr = 0

If Range("page_orientation").Value = "Landscape" And Range("page_size").Value = "Letter" Then
    btm = 690
ElseIf Range("page_orientation").Value = "Landscape" And Range("page_size").Value = "Tabloid" Then
    btm = 880
ElseIf Range("page_orientation").Value = "Portrait" And Range("page_size").Value = "Letter" Then
    btm = 690
End If

For x = 1 To lastrw Step 1
    If Worksheets("nqParts").Range("A" & x).top > btm Then
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
    Worksheets("nqParts").Range("A1:F" & lastrw).Copy
Else
    Worksheets("nqParts").Range("A1:F" & firstnr - 1).Copy
End If

With Worksheets("N+Q")
    Set nqPic1 = .Pictures.Paste(Link:=True)
    nqPic1.Name = "nqPic1"
    nqPic1.Left = .Range("A8").Left
    nqPic1.top = .Range("A8").top
End With

Dim nqPic2 As Picture

If firstnr <> "0" Then
Worksheets("nqParts").Range("A" & firstnr & ":F" & lastrw).Copy

With Worksheets("N+Q")
    Set nqPic1 = .Pictures.Paste(Link:=True)
    nqPic1.Name = "nqPic2"
    nqPic1.Left = .Range("E8").Left
    nqPic1.top = .Range("E8").top
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

If Range("page_size").Value = "Letter" And Range("page_orientation").Value = "Landscape" Then
    Sheets("N+Q").Columns("C").ColumnWidth = 85
    Sheets("N+Q").Columns("E").ColumnWidth = 85
    Sheets("N+Q").Rows("8:10").RowHeight = 235
ElseIf Range("page_size").Value = "Tabloid" And Range("page_orientation").Value = "Landscape" Then
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

Sub BIM()
Dim bimArray As Variant
Dim i As Long

bimArray = Array("BIM-1", "BIM-2", "BIM-3", "BIM-4", "BIM-5", "BIM-6", "BIM-7", "BIM-8")

For i = 0 To UBound(bimArray)
    If i < Range("bim").Value Then
        Sheets(bimArray(i)).Visible = True
        
        If Range("page_size").Value = "Letter" And Range("page_orientation").Value = "Landscape" Then

        ElseIf Range("page_size").Value = "Tabloid" And Range("page_orientation").Value = "Landscape" Then
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
    ElseIf Range("bim").Value = 0 Then
        If Sheets(bimArray(i)).Visible = True Then
            Sheets(bimArray(i)).Visible = False
        End If
    Else
        If Sheets(bimArray(i)).Visible = True Then
            Sheets(bimArray(i)).Visible = False
        End If
    End If
    
Next i

End Sub

Sub purgeImages()

Dim pic As Variant
Dim n As Integer
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
    For Each pic In ws.Pictures
        If pic.Name = "nqPic1" Or pic.Name = "nqPic2" Or pic.Name = "full_logo" Or pic.Name = "exec_1" Or pic.Name = "exec_2" Or pic.Name = "exec_3" Or pic.Name = "exec_4" Or pic.Name = "block_logo" Then
        Else
            pic.Delete
            n = n + 1
        End If
    Next
Next

MsgBox (n & " pictures were deleted successfully.")

End Sub


