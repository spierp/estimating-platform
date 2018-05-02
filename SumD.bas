Attribute VB_Name = "SumD"
Sub variance(report As String)

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
lineitemNumber = WorksheetFunction.CountA(range("B12:B120"))

Worksheets(report).Activate

Cells.Select
Selection.EntireColumn.Hidden = False
Selection.EntireRow.Hidden = False
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
    bottomRange = lineitemNumber + 11
    Worksheets(assocSum).range("B12:C" & bottomRange).Copy
    Sheets(report).range("B12").PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
pb.AddProgress 5

If range("var_show_comments").Value = "No" Then
    Columns("O").EntireColumn.Hidden = True
End If

If range("var_show_prim_div").Value = "No" Then
    Columns("E").EntireColumn.Hidden = True
    Columns("I").EntireColumn.Hidden = True
    Columns("M").EntireColumn.Hidden = True
End If

If range("var_show_sec_div").Value = "No" Then
    Columns("F").EntireColumn.Hidden = True
    Columns("J").EntireColumn.Hidden = True
    Columns("N").EntireColumn.Hidden = True
End If

End Sub
