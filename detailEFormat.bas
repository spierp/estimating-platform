Attribute VB_Name = "detailEFormat"
Sub tableFormat()
pb.Repaint
pb.AddCaption "Formatting Table..."
Rows(6).HorizontalAlignment = xlCenter

'WRAP TEXT
Columns("L:L").Select
    With Selection
        .WrapText = True
    End With
    
'INDENT DETAIL

Dim dat As Variant
Dim rng As range
Dim i As Long

Set rng = range("$J$7", Cells(Rows.count, "L").End(xlUp)).Cells
dat = rng.Value
For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 1) = "S" Then
    ElseIf dat(i, 1) = "H" Then
    ElseIf dat(i, 1) = "*" Then
        dat(i, 3) = "     " & dat(i, 3)
    ElseIf dat(i, 3) <> "" Then
        dat(i, 3) = "     " & dat(i, 3)
    End If
Next
rng.Value = dat

'Range("L7").Select
'Do Until WorksheetFunction.CountA(ActiveCell.Range("A1:A10")) < 1
'        If ActiveCell.Offset(0, -2).Value = "S" Then
''            ActiveCell.Font.Underline = True
'        ElseIf ActiveCell.Offset(0, -2).Value = "H" Then
''            ActiveCell.Font.Underline = True
'        ElseIf ActiveCell.Offset(0, -2).Value = "*" Then
'            ActiveCell.Font.Italic = True
'            ActiveCell.InsertIndent 2
'        Else
'            ActiveCell.InsertIndent 2
'        End If
'    ActiveCell.Offset(1, 0).Range("A1").Select
'Loop

'Format Table
With range("B6:C6").Font
    .ThemeColor = xlThemeColorAccent5
    .TintAndShade = -0.249977111117893
End With

range("A8").HorizontalAlignment = xlLeft

range("A6").CurrentRegion.Select
Selection.VerticalAlignment = xlVAlignCenter

Dim tbl As ListObject
Set tbl = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
tbl.TableStyle = "lineitem"
tbl.Name = ActiveSheet.Name & "Table"

    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""*Total*"",$A6))"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
'    Selection.FormatConditions(1).Borders(xlBottom).LineStyle = xlContinuous
'    With Selection.FormatConditions(1).Interior
'        .PatternColorIndex = xlAutomatic
'        .Color = RGB(238, 245, 252)
'    End With
    With Selection.FormatConditions(1).Font
        .Color = RGB(48, 84, 150)
        .TintAndShade = 0
        .Italic = True
        .Bold = True
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""*Total*"",$B6))"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
'    Selection.FormatConditions(1).Borders(xlBottom).LineStyle = xlContinuous
'    With Selection.FormatConditions(1).Interior
'        .PatternColorIndex = xlAutomatic
'        .Color = RGB(238, 245, 252)
'    End With
    With Selection.FormatConditions(1).Font
        .Color = RGB(48, 84, 150)
        .TintAndShade = 0
        .Italic = True
        .Bold = True
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""*Total*"",$C6))"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
'    Selection.FormatConditions(1).Borders(xlBottom).LineStyle = xlContinuous
'    With Selection.FormatConditions(1).Interior
'        .PatternColorIndex = xlAutomatic
'        .Color = RGB(238, 245, 252)
'    End With
    With Selection.FormatConditions(1).Font
        .Color = RGB(48, 84, 150)
        .TintAndShade = 0
        .Italic = True
        .Bold = True
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
'subtotal line item
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ISFORMULA(A6)"
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
    Selection.FormatConditions(1).Borders(xlTop).LineStyle = xlContinuous
'    With Selection.FormatConditions(1).Interior
'        .PatternColorIndex = xlAutomatic
'        .Color = RGB(238, 245, 252)
'    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$G6<>"""""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
'    With Selection.FormatConditions(1).Font
'        .Color = -11250480
'        .TintAndShade = 0
'    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$G6<>"""""
    Selection.FormatConditions(Selection.FormatConditions.count).SetFirstPriority
'    With Selection.FormatConditions(1).Font
'        .Color = -11250480
'        .TintAndShade = 0
'    End With
    Selection.FormatConditions(1).StopIfTrue = False

'    Columns("A:A").Select
'    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
'        "=LEN(TRIM(A1))=0"
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    Selection.FormatConditions(1).Borders(xlLeft).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlRight).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlTop).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlBottom).LineStyle = xlNone
'    With Selection.FormatConditions(1).Interior
'        .Color = RGB(255, 255, 255)
'        .TintAndShade = 0
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False
'
'    Columns("B:B").Select
'    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
'        "=AND(LEN(TRIM(B1))=0,LEN(TRIM(A1))=0)"
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    Selection.FormatConditions(1).Borders(xlLeft).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlRight).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlTop).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlBottom).LineStyle = xlNone
'    With Selection.FormatConditions(1).Interior
'        .Color = RGB(255, 255, 255)
'        .TintAndShade = 0
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False
'
'    Columns("C:C").Select
'    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
'        "=AND(LEN(TRIM(C1))=0,LEN(TRIM(B1))=0,LEN(TRIM(A1))=0)"
'    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
'    Selection.FormatConditions(1).Borders(xlLeft).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlRight).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlTop).LineStyle = xlNone
'    Selection.FormatConditions(1).Borders(xlBottom).LineStyle = xlNone
'    With Selection.FormatConditions(1).Interior
'        .Color = RGB(255, 255, 255)
'        .TintAndShade = 0
'    End With
'    Selection.FormatConditions(1).StopIfTrue = False
pb.AddProgress 3

' CREATE SPACER LINES
pb.AddCaption "Creating spacer lines between Level 1 sections"
range("A7").Select
Do Until WorksheetFunction.CountA(ActiveCell.Offset(0, 15).range("A1:A30")) < 1
    If ActiveCell.Value Like "*Total*" = True Then
        If ActiveSheet.Name = "brkDetail" Or ActiveSheet.Name = "altDetail" Then
            ActiveCell.RowHeight = 22
            ActiveSheet.HPageBreaks.Add Before:=Rows(ActiveCell.Offset(1, 0).Row)
            ActiveCell.Offset(1, 0).range("A1").Select
        Else: ActiveCell.RowHeight = 22
            ActiveCell.Offset(1, 0).range("A1").Select
            Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
            ActiveCell.RowHeight = 22
        End If
        
    End If
ActiveCell.Offset(1, 0).range("A1").Select
Loop
pb.AddProgress 5



'Range("B7").Select
'Do Until WorksheetFunction.CountA(ActiveCell.Offset(0, 14).Range("A1:A30")) < 1
'    If ActiveCell.Value Like "*Total*" = True Then
'        ActiveCell.Offset(1, 0).Range("A1").Select
'        ActiveCell.RowHeight = 18
'        Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
'        ActiveCell.RowHeight = 18
'    End If
'ActiveCell.Offset(1, 0).Range("A1").Select
'Loop
'
'Range("C7").Select
'Do Until WorksheetFunction.CountA(ActiveCell.Offset(0, 13).Range("A1:A30")) < 1
'    If ActiveCell.Value Like "*Total*" = True Then
'        ActiveCell.Offset(1, 0).Range("A1").Select
'        ActiveCell.RowHeight = 14
'        Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
'        ActiveCell.RowHeight = 14
'    End If
'ActiveCell.Offset(1, 0).Range("A1").Select
'Loop


End Sub
