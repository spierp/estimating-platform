Attribute VB_Name = "qto"
Public qtoForm As qtouserform
Public qtoDog As ProgressBarHotdog
Public wbThis As String
Public qtowb As String
Public zonecount As Integer
Public zonemode As Boolean
Public fatalerror As Boolean
Option Compare Text

Sub qtoworkbookselect()
workbookselectuserformQTO.Show
End Sub

Public Sub dynamo(qtowb As String)

Dim lastrow As Integer, qtorng As Range, qtodat As Variant, i As Integer, x As Integer, dyndat As Variant, dynrng As Range

fatalerror = False

wbThis = ActiveWorkbook.Name
qtowb = qtowb

Application.Workbooks(qtowb).Worksheets("dynamo-export").Activate
lastrow = Cells(Rows.Count, 6).End(xlUp).Row

''---START TAG STUFF
'Set dynrng = Worksheets("dynamo-export").Range("E2:F" & lastrow).Cells
'
'dyndat = dynrng.Value
'
'For i = LBound(dyndat, 1) To UBound(dyndat, 1)
'    If dyndat(i, 1) <> "" Then
'        dyndat(i, 2) = dyndat(i, 1) & " - " & dyndat(i, 2)
'    End If
'Next
'
'dynrng.Value = dyndat
'
'Columns(5).EntireColumn.Delete
''---END TAG STUFF

'CREATE NEW TABS
Dim ws As Worksheet
For Each ws In Worksheets
    If ws.Name = "QTO" Then
        Application.DisplayAlerts = False
        Sheets("QTO").Delete
        Application.DisplayAlerts = True
    End If
Next
    Sheets.Add Type:=xlWorksheet
    ActiveSheet.Name = "QTO"
    
Application.Workbooks(wbThis).Worksheets("Dashboard").Range("F23:Q23").Copy
Sheets("QTO").Range("H1").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False

'COPY UNIQUE VALUES
Worksheets("dynamo-export").Activate
Worksheets("dynamo-export").Range("C1:G" & lastrow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("QTO").Range("B1"), Unique:=True

'SORT UNIQUE VALUES
ActiveWorkbook.Worksheets("QTO").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("QTO").Sort.SortFields.Add Key:=Range("C2:C1000") _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("QTO").Sort.SortFields.Add Key:=Range("B2:B1000") _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("QTO").Sort.SortFields.Add Key:=Range("D2:D1000") _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
ActiveWorkbook.Worksheets("QTO").Sort.SortFields.Add Key:=Range("E2:E1000") _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("QTO").Sort
    .SetRange Range("B2:E1000")
    .Header = xlGuess
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'!!!!!!!!!!!!!!TEMPORARY - DELETE GROUP!!!!!!!!!!!!!!
ActiveWorkbook.Worksheets("QTO").Columns(4).EntireColumn.Delete

'SETUP COLUMNS
If zonemode = False Then
    If Application.WorksheetFunction.CountBlank(Range("B2:B" & lastrow)) > 0 Then
        MsgBox ("User Error:  I can't sort by 'Level' when your take-off has line items with blank Level data.  Crash and burn...")
        fatalerror = True
        Exit Sub
    Else
        Worksheets("dynamo-export").Range("B:B").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("QTO").Range("A1"), Unique:=True
        'Sort
        Worksheets("QTO").Sort.SortFields.Clear
        Worksheets("QTO").Sort.SortFields.Add Key:=Range("A1:A15"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
            xlSortNormal
        With Worksheets("QTO").Sort
            .SetRange Range("A1:A15")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        
    End If
Else:
    If Application.WorksheetFunction.CountBlank(Range("A2:A" & lastrow)) > 0 Then
        MsgBox ("User Error:  I can't sort by 'zone' when your take-off has line items with blank zones.  Crash and burn...")
        fatalerror = True
        Exit Sub
    Else
        Worksheets("dynamo-export").Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("QTO").Range("A1"), Unique:=True
    End If
End If

Sheets("QTO").Range("A2:A15").Copy
Sheets("QTO").Activate
Range("G2").PasteSpecial _
    Paste:=xlPasteValuesAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=True

If zonemode = True Then

    For x = 7 To 18 Step 1
        For i = 7 To 18 Step 1
            If Cells(1, x).Value = Cells(2, i).Value And Cells(1, x) <> "" Then
             Cells(5, x).Value = "matched"
            End If
        Next i
    Next x
      
    For x = 7 To 18 Step 1
        If Cells(1, x) <> "" And Cells(5, x) <> "matched" Then
            MsgBox ("User Error:  I can't align your data when the Revit zone tags don't match the estimate column names...  Please adjust in Revit or on your estimate Dashboard and try again.")
            fatalerror = True
            Exit Sub
        End If
    Next x

End If

Sheets("QTO").Range("A2:A15").ClearContents

zonecount = WorksheetFunction.CountA(Range("G1:R1"))

Cells(1, 1).Value = "STATUS"
Cells(1, 2).Value = "UNIFORMAT"
Cells(1, 3).Value = "CONTRACT ITEM"
Cells(1, 4).Value = "LINE ITEM"
Cells(1, 5).Value = "UNIT"
Cells(1, 6).Value = "TOTAL"

'RUN FORMULAS
Range("A1").CurrentRegion.Select
Set qtorng = Selection.Cells
qtodat = qtorng.Value

If zonemode = False Then
    For i = LBound(qtodat, 1) + 1 To UBound(qtodat, 1)
        qtodat(i, 6) = "=sum(G" & i & ":R" & i & ")"
        qtodat(i, 7) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!G1)"
        If zonecount > 1 Then
        qtodat(i, 8) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!H1)"
        End If
        If zonecount > 2 Then
        qtodat(i, 9) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!I1)"
        End If
        If zonecount > 3 Then
        qtodat(i, 10) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!J1)"
        End If
        If zonecount > 4 Then
        qtodat(i, 11) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!K1)"
        End If
        If zonecount > 5 Then
        qtodat(i, 12) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!L1)"
        End If
        If zonecount > 6 Then
        qtodat(i, 13) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!M1)"
        End If
        If zonecount > 7 Then
        qtodat(i, 14) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!N1)"
        End If
        If zonecount > 8 Then
        qtodat(i, 15) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!O1)"
        End If
        If zonecount > 9 Then
        qtodat(i, 16) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!P1)"
        End If
        If zonecount > 10 Then
        qtodat(i, 17) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!Q1)"
        End If
        If zonecount > 11 Then
        qtodat(i, 18) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!B:B,QTO!R1)"
        End If
    Next
Else
    For i = LBound(qtodat, 1) + 1 To UBound(qtodat, 1)
        qtodat(i, 6) = "=sum(G" & i & ":R" & i & ")"
        qtodat(i, 7) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!G1)"
        If zonecount > 1 Then
        qtodat(i, 8) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!H1)"
        End If
        If zonecount > 2 Then
        qtodat(i, 9) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!I1)"
        End If
        If zonecount > 3 Then
        qtodat(i, 10) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!J1)"
        End If
        If zonecount > 4 Then
        qtodat(i, 11) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!K1)"
        End If
        If zonecount > 5 Then
        qtodat(i, 12) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!L1)"
        End If
        If zonecount > 6 Then
        qtodat(i, 13) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!M1)"
        End If
        If zonecount > 7 Then
        qtodat(i, 14) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!N1)"
        End If
        If zonecount > 8 Then
        qtodat(i, 15) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!O1)"
        End If
        If zonecount > 9 Then
        qtodat(i, 16) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!P1)"
        End If
        If zonecount > 10 Then
        qtodat(i, 17) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!Q1)"
        End If
        If zonecount > 11 Then
        qtodat(i, 18) = "=SUMIFS('dynamo-export'!H:H,'dynamo-export'!C:C,QTO!B" & i & ",'dynamo-export'!D:D,QTO!C" & i & ",'dynamo-export'!F:F,QTO!D" & i & ",'dynamo-export'!A:A,QTO!R1)"
        End If
    Next
End If


qtorng.Value = qtodat

End Sub
