Attribute VB_Name = "detailADataTransfer"
Sub DetailTransfer(report As String)
Attribute DetailTransfer.VB_ProcData.VB_Invoke_Func = " \n14"
pb.Repaint

'CREATE NEW TAB
Dim ws As Worksheet
For Each ws In Worksheets
    If ws.Name = report Then
        Application.DisplayAlerts = False
        Sheets(report).Delete
        Application.DisplayAlerts = True
    End If
Next
Sheets.Add Type:=xlWorksheet
ActiveSheet.Name = report
    
'COPY DATA
pb.AddCaption "Copying data... "
pb.AddProgress 5
    Worksheets("Data").ListObjects("dataTable").AutoFilter.ShowAllData
    Worksheets("Data").ListObjects("dataTable").Range.Copy
    Worksheets(report).Range("A6").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False
    
    Worksheets("Data").Range("Q4:AN4").Copy
    Worksheets(report).Range("Q6").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
'FORMAT THIS
With Columns("Q:Q").Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = 0
    .Weight = xlMedium
End With
With Columns("AC:AC").Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ThemeColor = 1
    .TintAndShade = 0
    .Weight = xlMedium
End With
        
'DELETE UNUSED COLUMNS
Dim x As Long
For x = 40 To 16 Step -1
    If Cells(6, x).Value = "0" Or Cells(6, x).Value = "0_EXT" Then
        Columns(x).Delete
    End If
Next x
    
'FORMAT CODING
    Columns("H:J").Replace What:="_", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("J:J").Replace What:=".", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
'COLUMN RELOCATION
If report = "tradeDetail" Then
    Columns("J:J").Cut
    Columns("H:H").Insert Shift:=xlToRight
ElseIf report = "brkDetail" Then
    Columns("J:J").Cut
    Columns("H:H").Insert Shift:=xlToRight
    Columns("C:C").Cut
    Columns("H:H").Insert Shift:=xlToRight
    Columns("J:J").Cut
    Columns("A:A").Insert Shift:=xlToRight
ElseIf report = "altDetail" Then
    Columns("J:J").Cut
    Columns("H:H").Insert Shift:=xlToRight
    Columns("D:D").Cut
    Columns("H:H").Insert Shift:=xlToRight
    Columns("J:J").Cut
    Columns("A:A").Insert Shift:=xlToRight
End If

'SORT DATA
Range("H6").CurrentRegion.Select

    If WorksheetFunction.CountA(Range("J7:J50")) > 2 Then
        Selection.Sort Key1:=Range("J6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If

If report = "uniDetail" Then
    If WorksheetFunction.CountA(Range("I7:I50")) < 2 Then
        Selection.Sort Key1:=Range("H6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(Range("I7:I50")) > 2 Then
        Selection.Sort Key1:=Range("I6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If

ElseIf report = "tradeDetail" Then
    If WorksheetFunction.CountA(Range("J7:J50")) < 2 Then
        Selection.Sort Key1:=Range("I6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(Range("H7:H50")) > 2 Then
        Selection.Sort Key1:=Range("H6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If

ElseIf report = "brkDetail" Then
    If WorksheetFunction.CountA(Range("J7:J50")) > 2 Then
        Selection.Sort Key1:=Range("J6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(Range("I7:I50")) > 2 Then
        Selection.Sort Key1:=Range("I6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
        Selection.Sort Key1:=Range("H6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
          

ElseIf report = "altDetail" Then
    If WorksheetFunction.CountA(Range("J7:J50")) > 2 Then
        Selection.Sort Key1:=Range("J6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(Range("I7:I50")) > 2 Then
        Selection.Sort Key1:=Range("I6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
        Selection.Sort Key1:=Range("H6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
End If

'COLUMN HEADINGS
Range("L6").Value = "LINE ITEM"

If report = "uniDetail" Then
   Range("H6").Value = "CODE"
   Range("I6").Value = "UNI3/4"
   Range("J6").Value = "CI"
ElseIf report = "tradeDetail" Then
   Range("H6").Value = "CODE"
   Range("I6").Value = "UNI2"
   Range("J6").Value = "UNI3/4"
ElseIf repoty = "brkDetail" Then
   Range("H6").Value = "BRK"
   Range("I6").Value = "CI"
   Range("J6").Value = "UNI"
ElseIf repoty = "altDetail" Then
   Range("H6").Value = "ALT"
   Range("I6").Value = "CI"
   Range("J6").Value = "UNI"
End If

'REMOVE EXTRA LINES
If report = "brkDetail" Or report = "altDetail" Then
    Dim toprow As Integer
    Dim bottomrow As Integer
    
    bottomrow = Cells(Rows.Count, 16).End(xlUp).Row
    toprow = Cells(Rows.Count, 8).End(xlUp).Row + 1
    
    Rows(toprow & ":" & bottomrow).EntireRow.Delete
    pb.AddProgress 4
    
    'REMOVE HASHTAGS ON ALT SUMMARY
    If report = "altDetail" Then
        Dim rng As Range, dat As Variant, i As Integer
        Dim s As String, indexOfDollar As String
        
        Set rng = Range("P8:AN" & bottomrow).Cells
        dat = rng.Value
        
        For i = LBound(dat, 1) To UBound(dat, 1)
            If dat(i, 1) Like "*#*" = True Then
                s = dat(i, 1)
                indexOfDollar = InStr(1, s, "$")
                dat(i, 1) = Right(s, Len(s) - indexOfDollar + 1)
            End If
            For y = 2 To 20
            If dat(i, y) Like "*#*" = True Then
                s = dat(i, y)
                indexOfDollar = InStr(1, s, "$")
                dat(i, y) = Right(s, Len(s) - indexOfDollar + 1)
            End If
            Next
        Next
        rng.Value = dat
    End If
End If

End Sub
