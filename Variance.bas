Attribute VB_Name = "Variance"
Sub workbookSelectVar()
workbookselectuserformvar.Show
End Sub

Sub generalVariance()

    
    If Range("trade_variance").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(10).Total > 3 Then
        Call progressIndicator_Begin("Trade Summary Variance Report")
        Sheets("tradeVar").Visible = True
        Call sumVariance("tradeVar")
        Call sumPageSetup
        Sheets("dashboard").Activate
        Call progressIndicator_End
    Else
        If Sheets("tradeVar").Visible = True Then
            Sheets("TradeVar").Visible = False
        End If
    End If
    If Range("uniformat_L2_variance").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(10).Total > 3 Then
        Call progressIndicator_Begin("Uniformat L2 Summary Variance Report")
        Sheets("uni2Var").Visible = True
        Call sumVariance("uni2Var")
        Call sumPageSetup
        Sheets("dashboard").Activate
        Call progressIndicator_End
    Else
        If Sheets("uni2Var").Visible = True Then
            Sheets("uni2Var").Visible = False
        End If
    End If
    If Range("uniformat_L34_variance").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(10).Total > 3 Then
        Call progressIndicator_Begin("Uniformat L4 Summary Variance Report")
        Sheets("uni34Var").Visible = True
        Call sumVariance("uni34Var")
        Call sumPageSetup
        Sheets("dashboard").Activate
        Call progressIndicator_End
    Else
        If Sheets("uni34Var").Visible = True Then
            Sheets("uni34Var").Visible = False
        End If
    End If
    
    If Range("detail_variance").Value = "Yes" And _
    Worksheets("Data").ListObjects("dataTable").ListColumns(10).Total > 3 Then
    Call workbookSelectVar
    End If
       
    
End Sub

Sub sumVariance(report As String)

Dim assocSum As String

If report = "tradeVar" Then
    assocSum = "tradeSum"
ElseIf report = "uni2Var" Then
    assocSum = "uni2Sum"
ElseIf report = "uni34Var" Then
    assocSum = "uni34Sum"
End If

'Creating correct number of rows..."
pb.AddCaption "Creating correct number of rows on variance tab..."

Worksheets(assocSum).Activate

Dim lineitemNumber As Integer
lineitemNumber = WorksheetFunction.CountA(Range("B12:B120"))

Worksheets(report).Activate

Cells.Select
Selection.EntireColumn.Hidden = False
Selection.EntireRow.Hidden = False
Range("B12").Select

    Do Until WorksheetFunction.CountA(Range("B12:B200")) = lineitemNumber
        If WorksheetFunction.CountA(Range("B12:B200")) < lineitemNumber Then
            ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
            Selection.Copy
            Selection.Insert Shift:=xlDown
            ActiveCell.Offset(1, 7).Range("A1").ClearContents
            Range("B12").Select
        ElseIf WorksheetFunction.CountA(Range("B12:B200")) > lineitemNumber Then
            ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
            ActiveCell.Activate
            Selection.Delete Shift:=xlDown
            Range("B12").Select
        End If
    Loop

pb.AddProgress 25

'Copy Data
pb.AddCaption "Copying data to Summary tab..."

    Dim bottomRange As String
    bottomRange = lineitemNumber + 11
    Worksheets(assocSum).Range("B12:C" & bottomRange).Copy
    Sheets(report).Range("B12").PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
pb.AddProgress 20
pb.AddCaption "Hiding columns that are not applicable..."

If Range("var_show_comments").Value = "No" Then
    Columns("O").EntireColumn.Hidden = True
End If

If Range("var_show_prim_div").Value = "No" Then
    Columns("E").EntireColumn.Hidden = True
    Columns("I").EntireColumn.Hidden = True
    Columns("N").EntireColumn.Hidden = True
End If

If Range("var_show_sec_div").Value = "No" Then
    Columns("F").EntireColumn.Hidden = True
    Columns("J").EntireColumn.Hidden = True
    Columns("O").EntireColumn.Hidden = True
End If

If Range("var_show_perc").Value = "No" Then
    Columns("M").EntireColumn.Hidden = True
End If

pb.AddProgress 25

pb.AddCaption "Hiding markup rows that are not applicable..."
ActiveSheet.Columns(3).Find("COST OF WORK - SUBTOTAL").Select
ActiveCell.Offset(2, 0).Select
    Do Until ActiveCell.Value = ""
        If ActiveCell.Offset(0, 1) = "0" Then
            ActiveCell.Rows("1:1").EntireRow.Select
            Selection.EntireRow.Hidden = True
            ActiveCell.Offset(0, 2).Select
        End If
        ActiveCell.Offset(1, 0).Select
    Loop
pb.AddProgress 10

End Sub

Sub detailVariance(comparablewb As String)

workbookselectuserformvar.Hide


Call progressIndicator_Begin("Detailed Variance Report")


Dim wbThis As Workbook
Set wbThis = ActiveWorkbook
Dim firstrw1 As Integer
Dim firstrw2 As Integer

'CREATE NEW TAB
Dim ws As Worksheet
For Each ws In Worksheets
    If ws.Name = "varDetail" Then
        Application.DisplayAlerts = False
        Sheets("varDetail").Delete
        Application.DisplayAlerts = True
    End If
Next
    Sheets.Add Type:=xlWorksheet
    ActiveSheet.Name = "varDetail"
    
'COPY PREVIOUS DATA
pb.AddCaption "Copying previous data... "
pb.AddProgress 3

Application.Workbooks(comparablewb).Worksheets("Data").Activate
Worksheets("Data").ListObjects("dataTable").AutoFilter.ShowAllData
Worksheets("Data").ListObjects("dataTable").Range.Copy
wbThis.Worksheets("Data").Activate
wbThis.Worksheets("varDetail").Range("A1").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False
    
Application.CutCopyMode = False

Sheets("varDetail").Activate
    
firstrw1 = Selection.Rows.Count + 1

Columns("A:D").EntireColumn.Delete
Columns("B:C").EntireColumn.Delete
Columns("K:AH").EntireColumn.Delete

'COPY NEW DATA
pb.AddCaption "Copying new data... "
pb.AddProgress 2

wbThis.Worksheets("Data").ListObjects("dataTable").AutoFilter.ShowAllData
wbThis.Worksheets("Data").ListObjects("dataTable").Range.Copy
wbThis.Worksheets("varDetail").Range("K1").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False

Application.CutCopyMode = False

Sheets("varDetail").Activate

firstrw2 = Selection.Rows.Count + 1

Columns("K:N").EntireColumn.Delete
Columns("L:M").EntireColumn.Delete
Columns("U:AR").EntireColumn.Delete

Range("K1:T" & firstrw2).Copy

wbThis.Worksheets("varDetail").Range("A" & firstrw1).PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False

Application.CutCopyMode = False

Columns("K:T").EntireColumn.Delete

Range("G1:J" & firstrw1 - 1).Cut wbThis.Worksheets("varDetail").Range("K1")

Application.CutCopyMode = False

Rows(firstrw1).EntireRow.Delete

Range("F:F").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Dim lastrow As Integer
lastrow = Cells(Rows.Count, 1).End(xlUp).Row + 5

Rows("1:5").EntireRow.Insert
Columns("G:I").EntireColumn.Insert

'TRACK CHANGES
pb.AddCaption "identifying changes... "
pb.AddProgress 3

Dim rng As Range
Dim dat As Variant
Dim i As Long
Dim j As Long

Set rng = Range("A7:Q" & lastrow).Cells
dat = rng.Value

'For i = LBound(dat, 1) To UBound(dat, 1)
'    If IsNumeric(dat(i, 13)) = False Then
'        dat(i, 13) = ""
'    End If
'    If IsNumeric(dat(i, 17)) = False Then
'        dat(i, 17) = ""
'    End If
'Next

'Define Description
For i = LBound(dat, 1) To UBound(dat, 1)
    For j = LBound(dat, 1) To UBound(dat, 1)
        If dat(j, 1) = dat(i, 1) And j <> i Then
            If (dat(i, 13)) <> (dat(j, 17)) Or (dat(j, 13)) <> (dat(i, 17)) Then
                dat(i, 1) = ""
                If dat(i, 13) = "" Then
                    dat(j, 14) = dat(i, 14)
                    dat(j, 15) = dat(i, 15)
                    dat(j, 16) = dat(i, 16)
                    dat(j, 17) = dat(i, 17)
                    If dat(j, 10) > dat(j, 14) And dat(j, 12) = dat(j, 16) Then
                        dat(j, 8) = "unit cost increased by " & Format(dat(j, 10) - dat(j, 14), "$#,##0.00") & " / " & dat(j, 11)
                    ElseIf dat(j, 10) < dat(j, 14) And dat(j, 12) = dat(j, 16) Then
                        dat(j, 8) = "unit cost decreased by " & Format(dat(j, 14) - dat(j, 10), "$#,##0.00") & " / " & dat(j, 11)
                    ElseIf dat(j, 12) > dat(j, 16) And dat(j, 10) = dat(j, 14) Then
                        dat(j, 8) = "quantity increased by " & Format(dat(j, 12) - dat(j, 16), "###,##0") & " " & dat(j, 11)
                    ElseIf dat(j, 12) < dat(j, 16) And dat(j, 10) = dat(j, 14) Then
                        dat(j, 8) = "quantity decreasesd by " & Format(dat(j, 16) - dat(j, 12), "###,##0") & " " & dat(j, 11)
                    ElseIf dat(j, 12) > dat(j, 16) And dat(j, 10) > dat(j, 14) Then
                        dat(j, 8) = "quantity increased by " & Format(dat(j, 12) - dat(j, 16), "###,##0") & " " & dat(j, 11) _
                        & " and unit cost increased by " & Format(dat(j, 10) - dat(j, 14), "$#,##0.00") & " / " & dat(j, 11)
                    ElseIf dat(j, 12) < dat(j, 16) And dat(j, 10) < dat(j, 14) Then
                        dat(j, 8) = "quantity decreased by " & Format(dat(j, 16) - dat(j, 12), "###,##0") & " " & dat(j, 11) _
                        & " and unit cost decreased by " & Format(dat(j, 14) - dat(j, 10), "$#,##0.00") & " / " & dat(j, 11)
                    ElseIf dat(j, 12) < dat(j, 16) And dat(j, 10) > dat(j, 14) Then
                        dat(j, 8) = "quantity decreasesd by " & Format(dat(j, 16) - dat(j, 12), "###,##0") & " " & dat(j, 11) _
                        & " and unit cost increased by " & Format(dat(j, 10) - dat(j, 14), "$#,##0.00") & " / " & dat(j, 11)
                    ElseIf dat(j, 12) > dat(j, 16) And dat(j, 10) < dat(j, 14) Then
                        dat(j, 8) = "quantity increased by " & Format(dat(j, 12) - dat(j, 16), "###,##0") & " " & dat(j, 11) _
                        & " and unit cost decreased by " & Format(dat(j, 14) - dat(j, 10), "$#,##0.00") & " / " & dat(j, 11)
                    End If
                End If
            Else
                dat(i, 1) = ""
                dat(j, 1) = ""
            End If
        Else
        End If
    Next
Next

'Calculate Delta

For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 13) = "" Then
    dat(i, 8) = "removed"
        If IsNumeric(dat(i, 17)) Then
            dat(i, 7) = dat(i, 17) * -1
        End If
    End If
    If dat(i, 17) = "" Or IsNumeric(dat(i, 17)) = False And dat(i, 13) <> "" Then
    dat(i, 8) = "added"
        If IsNumeric(dat(i, 13)) Then
            dat(i, 7) = dat(i, 13)
        End If
    End If
Next

For i = LBound(dat, 1) To UBound(dat, 1)
    If IsNumeric(dat(i, 17)) And IsNumeric(dat(i, 13)) Then
    dat(i, 7) = dat(i, 13) - dat(i, 17)
    End If
Next

Sheets("varDetail").Cells.Delete
[A1].Resize(UBound(dat), UBound(Application.Transpose(dat))) = dat
Range("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete

Rows("1:6").EntireRow.Insert

'FORMAT CODING
pb.AddCaption "identifying changes... "
pb.AddProgress 5

    Columns("B:D").Replace What:="_", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("D:D").Replace What:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

'CREATE HEADERS
pb.AddCaption "Creating Headings... "
pb.AddProgress 2

Range("A6").Value = "GUID"
Range("B6").Value = "UNI2"
Range("C6").Value = "UNI34"
Range("D6").Value = "CODE"
Range("E6").Value = "SPACE"
Range("F6").Value = "LINE ITEM"
Range("G6").Value = "DELTA"
Range("H6").Value = "DESCRIPTION"
Range("I6").Value = "COMMENTS"
Range("J6").Value = "N-U/P"
Range("K6").Value = "N-U"
Range("L6").Value = "N-QTY"
Range("M6").Value = "N-TOTAL"
Range("N6").Value = "P-U/P"
Range("O6").Value = "P-U"
Range("P6").Value = "P-QTY"
Range("Q6").Value = "P-TOTAL"

Columns("A:A").EntireColumn.Delete

Columns("C:C").Cut
Columns("A:A").Insert Shift:=xlToRight
Columns("D:D").EntireColumn.Hidden = True
Range("A1").Select

'SORT DATA
pb.AddCaption "Sorting Data... "
pb.AddProgress 3

    Range("H6").CurrentRegion.Select

    If WorksheetFunction.CountA(Range("C7:C50")) > 2 Then
        Selection.Sort Key1:=Range("C6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(Range("C7:C50")) < 2 Then
        Selection.Sort Key1:=Range("B6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(Range("A7:A50")) > 2 Then
        Selection.Sort Key1:=Range("A6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If

'FORMAT COLUMN WIDTHS
pb.AddProgress 3
pb.AddCaption "formatting column widths... "

    Columns("A:B").ColumnWidth = 5
    Columns("C:C").ColumnWidth = 1
    Columns("E:E").ColumnWidth = 55
    Columns("F:F").ColumnWidth = 16
    Columns("G:G").ColumnWidth = 55
    Columns("H:H").ColumnWidth = 35
    Columns("I:I").ColumnWidth = 11
    Columns("M:M").ColumnWidth = 11
    Columns("J:J").ColumnWidth = 6
    Columns("N:N").ColumnWidth = 6
    Columns("K:K").ColumnWidth = 11
    Columns("O:O").ColumnWidth = 11
    Columns("L:L").ColumnWidth = 13
    Columns("P:P").ColumnWidth = 13


'FORMAT FONT & ALIGNMENT
pb.AddProgress 3
pb.AddCaption "formatting fonts and alignment... "

    With Columns("L:L")
        .HorizontalAlignment = xlRight
    End With
    With Columns("P:P")
        .HorizontalAlignment = xlRight
    End With

    With Rows("6:6")
        .Font.Bold = True
        .RowHeight = 22
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With Columns("F:F")
            .Font.Bold = True
            .NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    End With
    With Columns("E:E")
        .WrapText = True
    End With

    With Columns("K:K")
        .Style = "Comma"
        .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    End With
    With Columns("O:O")
        .Style = "Comma"
        .NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    End With

    With Columns("L:L")
        .Style = "Currency"
        .NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    End With
    With Columns("P:P")
        .Style = "Currency"
        .NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    End With

'CREATE SUBTOTALS
pb.AddCaption "Calculating Subtotals..."
pb.AddProgress 3

Range("A6").CurrentRegion.Select
Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(6, 12, 16), Replace:=False, PageBreaks:=False, _
SummaryBelowData:=True
Range("A6").CurrentRegion.Select
Selection.ClearOutline

With Selection
    .VerticalAlignment = xlCenter
End With

lastrow = Range("A:A").Find(What:="Grand Total", After:=[A7], SearchDirection:=xlPrevious).Row
Range("A" & lastrow).Value = "Direct Work - Total Variance"

'CREATE AREA DIVISOR ON TOTALS
pb.AddCaption "Calculating area divisor on totals..."
pb.AddProgress 5

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

If ThisWorkbook.Names("detail_prim_div").RefersToRange(1, 1).Value = "Yes" Then

primDivUnit = ThisWorkbook.Names("prim_div_unit").RefersToRange(1, 1).Value
primDivQty = ThisWorkbook.Names("prim_div_qty").RefersToRange(1, 1).Value

For x = 7 To lastrow Step 1
    If IsNumeric(Cells(x, 6).Value) = True And Cells(x, 1).Font.Bold = True Or _
    IsNumeric(Cells(x, 6).Value) = True And Cells(x, 2).Font.Bold = True Then
        Cells(x, 7).Value = "$" & Round(Cells(x, 6).Value / primDivQty, 2) & " / " & primDivUnit
        Cells(x, 7).HorizontalAlignment = xlLeft
        Cells(x, 7).IndentLevel = 5
        Rows(x).RowHeight = 18
        Rows(x).Font.Bold = True
    End If
Next x
End If

'CREATE HEADINGS
pb.AddCaption "Creating Level 2 Headings..."
pb.AddProgress 10

lastrow = Cells(Rows.Count, 6).End(xlUp).Row

For x = lastrow To 7 Step -1
    If Cells(x, 1).Font.Bold = False _
    And Cells(x, 1).Value <> Cells(x - 1, 1).Value _
    And Cells(x - 1, 1).Value <> "" _
    And Cells(x, 1).Value <> "" Then
        Rows(x).EntireRow.Insert
        lastrow = lastrow + 1
        Rows(x).RowHeight = 22
        Rows(x).Font.Bold = True
        Cells(x, 1).Value = Cells(x + 1, 1).Value
    End If

Next x
pb.AddProgress 5

pb.AddCaption "Creating Level 2 Headings... (this might take a longer minute)"
For x = lastrow To 7 Step -1
    If Cells(x, 2).Font.Bold = False _
    And Cells(x, 2).Value <> Cells(x - 1, 2).Value _
    And Cells(x, 2).Value <> Cells(x - 2, 2).Value _
    And Cells(x, 2).Value <> "" Then
        Rows(x).EntireRow.Insert
        lastrow = lastrow + 1
        Rows(x).RowHeight = 18
        Rows(x).Font.Bold = True
        Cells(x, 2).Value = Cells(x + 1, 2).Value
    End If
Next x
pb.AddProgress 5

pb.AddCaption "Creating Level 3 Headings... (this is the longest minute)"
For x = lastrow To 7 Step -1
    If Cells(x, 3).Font.Bold = False _
    And Cells(x, 3).Value <> Cells(x - 1, 3).Value _
    And Cells(x, 3).Value <> "" Then
        Rows(x).EntireRow.Insert
        lastrow = lastrow + 1
        Rows(x).RowHeight = 14
        Rows(x).Font.Bold = True
        Cells(x, 3).Value = Cells(x + 1, 3).Value
    End If

Next x
pb.AddProgress 10

'DELETE EXTRA CODES
pb.AddCaption "Deleting extra code references..."

'Dim dat As Variant
'Dim rng As range
'Dim i As Long

Set rng = Range("$A$7", Cells(Rows.Count, "C").End(xlUp)).Cells
dat = rng.Value
For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 1) <> "" And dat(i, 2) <> "" And dat(i, 3) <> "" Then
    dat(i, 1) = ""
    dat(i, 2) = ""
    dat(i, 3) = ""
    End If
Next
rng.Value = dat

Erase dat
pb.AddProgress 5

'ADD ROW NUMBERS
pb.AddCaption "Adding row numbers..."

Dim lineCount As Integer
lineCount = 0

Set rng = Range("$A$7", Cells(Rows.Count, "E").End(xlUp)).Cells
dat = rng.Value
For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 5) <> "" Then
    lineCount = lineCount + 1
    dat(i, 1) = lineCount
    End If
Next
rng.Value = dat
pb.AddProgress 5

'FORMAT TABLE
Range("A6").CurrentRegion.Select
With Range("B6:C6").Font
    .ThemeColor = xlThemeColorAccent5
    .TintAndShade = -0.249977111117893
End With

Dim tbl As ListObject
Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
tbl.TableStyle = "lineitem"
tbl.Name = ActiveSheet.Name & "Table"

    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""*Total*"",$A6))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = RGB(48, 84, 150)
        .TintAndShade = 0
        .Italic = True
        .Bold = True
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""*Total*"",$B6))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = RGB(48, 84, 150)
        .TintAndShade = 0
        .Italic = True
        .Bold = True
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""*Total*"",$C6))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = RGB(48, 84, 150)
        .TintAndShade = 0
        .Italic = True
        .Bold = True
    End With
    Selection.FormatConditions(1).StopIfTrue = False

'DIRECT WORK BOTTOM
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""*Direct Work*"",$A6))"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Interior.Color = RGB(255, 242, 204)
    Selection.FormatConditions(1).StopIfTrue = False

Columns("G").InsertIndent 1

'SUBTOTAL LINE ITEM
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISFORMULA(A6)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).Borders(xlTop).LineStyle = xlContinuous
    Selection.FormatConditions(1).StopIfTrue = False

' CREATE SPACER LINES
pb.AddCaption "Creating spacer lines between Level 1 sections"

Range("A7").Select
Do Until WorksheetFunction.CountA(ActiveCell.Offset(0, 6).Range("A1:A30")) < 1
    If ActiveCell.Value Like "*Total*" = True Then
        ActiveCell.RowHeight = 22
            ActiveCell.Offset(1, 0).Range("A1").Select
            Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
            ActiveCell.RowHeight = 22
    End If
ActiveCell.Offset(1, 0).Range("A1").Select
Loop

pb.AddProgress 5

'Add Header Info
pb.AddCaption "Creating Page Header..."

'Insert Logo
Worksheets("dashboard").Shapes("full_logo").Copy
Range("A1").Select
ActiveSheet.Paste
Selection.ShapeRange.ScaleHeight 0.5715085068, msoFalse, msoScaleFromTopLeft

'Center Header
Range("C1").Value = UCase(ThisWorkbook.Names("project_name").RefersToRange(1, 1).Value)
Range("C2").Value = UCase(ThisWorkbook.Names("client_name").RefersToRange(1, 1).Value)
Range("C3").Value = UCase(ThisWorkbook.Names("estimate_name").RefersToRange(1, 1).Value)

Range("I5").Value = "NEW"
Range("I5:L5").Select

    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .Bold = True
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
    End With

Range("M5").Value = "PREVIOUS"
Range("M5:P5").Select

    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .Bold = True
    End With

Range("C1:P4").Select

    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
    With Selection.Font
        .Bold = True
    End With

Range("C2").Font.Underline = True

'Right Header
Range("P1:P3").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With

Range("P1").Value = "DETAILED VARIANCE REPORT"
Range("P2").Value = ThisWorkbook.Names("estimate_date").RefersToRange(1, 1).Value
Range("P2").NumberFormat = "mm/dd/yyyy"
Range("P3").Value = 1
Range("P3").Font.Color = RGB(255, 255, 255)

Rows("1").Select
    With Selection.Font
        .Size = 11
    End With

Rows("7").Select
ActiveWindow.FreezePanes = True
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.RowHeight = 12

Range("A8:C10").HorizontalAlignment = xlLeft

Columns("A:B").ColumnWidth = 5
Columns("C").ColumnWidth = 1

pb.AddProgress 3

lastrow = Range("A:A").Find(What:="Direct Work - Total Variance", After:=[A7], SearchDirection:=xlPrevious).Row
Range("I5:L" & lastrow).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
Range("M5:P" & lastrow).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

Range("A1").Select

'CONFIGURING PRINT SETUP
pb.AddCaption "Configuring Print Setup..."

Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$7"
        .PrintTitleColumns = ""
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "Page &P of &N"
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
        .CenterVertically = False
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
        .FitToPagesTall = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.text = ""
        .EvenPage.CenterHeader.text = ""
        .EvenPage.RightHeader.text = ""
        .EvenPage.LeftFooter.text = ""
        .EvenPage.CenterFooter.text = ""
        .EvenPage.RightFooter.text = ""
        .FirstPage.LeftHeader.text = ""
        .FirstPage.CenterHeader.text = ""
        .FirstPage.RightHeader.text = ""
        .FirstPage.LeftFooter.text = ""
        .FirstPage.CenterFooter.text = ""
        .FirstPage.RightFooter.text = ""
    End With
    Application.PrintCommunication = True

Sheets("varDetail").Move After:=Worksheets(Worksheets.Count)
With ActiveWorkbook.Sheets("tradeDetail").Tab
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = 0.399975585192419
End With
        
Sheets("dashboard").Activate
Call progressIndicator_End
End Sub
