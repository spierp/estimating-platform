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
    Worksheets("Data").ListObjects("dataTable").range.Copy
    Worksheets(report).range("A6").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False
    
    Worksheets("Data").range("Q4:AN4").Copy
    Worksheets(report).range("Q6").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=True, Transpose:=False
        
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
    range("H6").CurrentRegion.Select

    If WorksheetFunction.CountA(range("J7:J50")) > 2 Then
        Selection.Sort Key1:=range("J6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    
If report = "uniDetail" Then
    If WorksheetFunction.CountA(range("I7:I50")) < 2 Then
        Selection.Sort Key1:=range("H6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(range("I7:I50")) > 2 Then
        Selection.Sort Key1:=range("I6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If

ElseIf report = "tradeDetail" Then
    If WorksheetFunction.CountA(range("J7:J50")) < 2 Then
        Selection.Sort Key1:=range("I6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(range("H7:H50")) > 2 Then
        Selection.Sort Key1:=range("H6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    
ElseIf report = "brkDetail" Then
    If WorksheetFunction.CountA(range("I7:I50")) > 2 Then
        Selection.Sort Key1:=range("I6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(range("H7:H50")) > 2 Then
        Selection.Sort Key1:=range("H6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
ElseIf report = "altDetail" Then
    If WorksheetFunction.CountA(range("J7:J50")) > 2 Then
        Selection.Sort Key1:=range("J6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(range("I7:I50")) > 2 Then
        Selection.Sort Key1:=range("I6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
    If WorksheetFunction.CountA(range("H7:H50")) > 2 Then
        Selection.Sort Key1:=range("H6"), Order1:=xlAscending, Header:=xlGuess, _
            OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
            DataOption1:=xlSortNormal
    End If
End If
    
'COLUMN HEADINGS
range("L6").Value = "LINE ITEM"
If report = "uniDetail" Then
   range("H6").Value = "CODE"
   range("I6").Value = "UNI3/4"
   range("J6").Value = "CI"
ElseIf report = "tradeDetail" Then
   range("H6").Value = "CODE"
   range("I6").Value = "UNI2"
   range("J6").Value = "UNI3/4"
ElseIf repoty = "brkDetail" Then
   range("H6").Value = "BRK"
   range("I6").Value = "CI"
   range("J6").Value = "UNI"
ElseIf repoty = "altDetail" Then
   range("H6").Value = "ALT"
   range("I6").Value = "CI"
   range("J6").Value = "UNI"
End If

'REMOVE EXTRA LINES
If report = "brkDetail" Or report = "altDetail" Then
Dim toprow As Integer
Dim bottomrow As Integer

bottomrow = Cells(Rows.count, 16).End(xlUp).Row
toprow = Cells(Rows.count, 8).End(xlUp).Row + 1

Rows(toprow & ":" & bottomrow).EntireRow.Delete
pb.AddProgress 4
End If

'REMOVE HASHTAGS
If report = "altDetail" Then
pb.AddCaption "Removing Hashtags... "
    range("P7").Select
    Do Until ActiveCell.range("A1") = ""
        If ActiveCell.Value Like "*#*" = True Then
            
            Dim s As String
            s = ActiveCell.Value

            Dim indexOfDollar As Integer
            indexOfDollar = InStr(1, s, "$")

            Dim finalString As String
            finalString = Right(s, Len(s) - indexOfDollar + 1)
            
            ActiveCell.Value = finalString
            
        End If
    ActiveCell.Offset(1, 0).range("A1").Select
    Loop
End If



End Sub
