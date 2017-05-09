Attribute VB_Name = "sumA"
Sub sumData(report As String)
Sheets("clipboard").Visible = True

'Set variables
Dim datacolumn As Integer
Dim clipcolumn As String
Dim sortcolumn As String
Dim toprange As String

If report = "tradeSum" Or report = "tradeVar" Then
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
Worksheets("Data").ListObjects("dataTable").ListColumns(datacolumn).range.Copy
Sheets("clipboard").range(clipcolumn & "1").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False
    
Worksheets("clipboard").Activate
    
pb.AddProgress 10
 
'Filter unique values
pb.AddCaption "Filtering for unique categories..."

Worksheets("clipboard").range(clipcolumn & ":" & clipcolumn).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("Clipboard").range(sortcolumn & "1"), Unique:=True

'Sort Contract Items
    Worksheets("clipboard").Sort.SortFields.Clear
    Worksheets("clipboard").Sort.SortFields.Add Key:=range( _
        sortcolumn & "1:" & sortcolumn & "500"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With Worksheets("Clipboard").Sort
        .SetRange range(sortcolumn & "1:" & sortcolumn & "500")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

pb.AddProgress 5

'Remove underscores
pb.AddCaption "Formatting text..."
Worksheets("clipboard").range(sortcolumn & ":" & sortcolumn).Replace What:="_", Replacement:=" ", LookAt:=xlPart, SearchOrder _
    :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
      

'Creating correct number of rows..."
pb.AddCaption "Creating correct number of rows on summary tab..."

Dim lineitemNumber As Integer
lineitemNumber = WorksheetFunction.CountA(range(sortcolumn & "3:" & sortcolumn & "100"))
Worksheets(report).Activate
range("B12").Select

    Do Until WorksheetFunction.CountA(range("B12:B200")) = lineitemNumber
        If WorksheetFunction.CountA(range("B12:B200")) < lineitemNumber Then
            ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
            Selection.Copy
            Selection.Insert Shift:=xlDown
            range("B12").Select
        ElseIf WorksheetFunction.CountA(range("B12:B200")) > lineitemNumber Then
            ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
            ActiveCell.Activate
            Selection.Delete Shift:=xlDown
            range("B12").Select
        End If
    Loop

pb.AddProgress 25

'Copy Data
pb.AddCaption "Copying data to Summary tab..."
    Dim bottomRange As String
    bottomRange = toprange & lineitemNumber + 2
    Worksheets("clipboard").range(bottomRange).Copy
    Sheets(report).range("B12").PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
pb.AddProgress 5
Sheets("clipboard").Visible = False
End Sub

