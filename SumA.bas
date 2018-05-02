Attribute VB_Name = "sumA"
Sub sumData(report As String)
Sheets("clipboard").Visible = True

'SET VARIABLES
Dim datacolumn As Integer
Dim clipcolumn As String
Dim sortcolumn As String
Dim toprange As String

If report = "tradeSum" Or report = "tradeVar" Or report = "brkSum" Or report = "altSum" Then
    datacolumn = "10"
    clipcolumn = "A"
    sortcolumn = "C"
    toprange = "D3:E"
ElseIf report = "uni2Sum" Then
    datacolumn = "8"
    clipcolumn = "G"
    sortcolumn = "I"
    toprange = "J3:K"
ElseIf report = "uni34Sum" Then
    datacolumn = "9"
    clipcolumn = "G"
    sortcolumn = "I"
    toprange = "J3:K"
End If

'Copy Contract Items to Clipboard
pb.AddCaption "Copying trade categories..."

Worksheets("clipboard").Columns("A:C").ClearContents
Worksheets("clipboard").Columns("G:I").ClearContents
Worksheets("Data").ListObjects("dataTable").AutoFilter.ShowAllData

Worksheets("Data").ListObjects("dataTable").ListColumns(datacolumn).Range.Copy
Sheets("clipboard").Range(clipcolumn & "1").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False
    
Worksheets("clipboard").Activate
    
pb.AddProgress 10
 
'Filter unique values
pb.AddCaption "Filtering for unique categories..."

Worksheets("clipboard").Range(clipcolumn & ":" & clipcolumn).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("Clipboard").Range(sortcolumn & "1"), Unique:=True

'Sort Contract Items
    Worksheets("clipboard").Sort.SortFields.Clear
    Worksheets("clipboard").Sort.SortFields.Add Key:=Range( _
        sortcolumn & "1:" & sortcolumn & "500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Worksheets("Clipboard").Sort
        .SetRange Range(sortcolumn & "1:" & sortcolumn & "500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

pb.AddProgress 5

'Remove underscores
pb.AddCaption "Formatting text..."
Worksheets("clipboard").Range(sortcolumn & ":" & sortcolumn).Replace What:="_", Replacement:=" ", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
      

'Creating correct number of rows..."
pb.AddCaption "Creating correct number of rows on summary tab..."

Dim lineitemNumber As Integer
lineitemNumber = WorksheetFunction.CountA(Range(sortcolumn & "3:" & sortcolumn & "100"))
Worksheets(report).Activate
Range("B12").Select

    Do Until WorksheetFunction.CountA(Range("B12:B200")) = lineitemNumber
        If WorksheetFunction.CountA(Range("B12:B200")) < lineitemNumber Then
            ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
            Selection.Copy
            Selection.Insert Shift:=xlDown
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
    bottomRange = toprange & lineitemNumber + 2
    Worksheets("clipboard").Range(bottomRange).Copy
    Sheets(report).Range("B12").PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
pb.AddProgress 5

End Sub

