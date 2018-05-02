Attribute VB_Name = "dataTable"
Public filt()

Sub dataSort()
Attribute dataSort.VB_ProcData.VB_Invoke_Func = " \n14"
    Range("dataTable").Select
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=Range("dataTable[CONTRACT ITEM]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=Range("dataTable[UNI L2]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=Range("dataTable[UNI  L3/L4]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
End Sub
Sub dataSortUNI()
    Range("dataTable").Select
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=Range("dataTable[UNI L2]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=Range("dataTable[UNI  L3/L4]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=Range("dataTable[CONTRACT ITEM]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
End Sub

Sub insert5rows()
Attribute insert5rows.VB_ProcData.VB_Invoke_Func = " \n14"
If ActiveSheet.Name = "Data" Then
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        Dim rw As Integer
        rw = ActiveCell.Row
        Rows(rw + 1 & ":" & rw + 5).Insert Shift:=xlDown
        Range("H" & rw + 1 & ":" & "J" & rw + 5).Value = Range("H" & rw & ":" & "J" & rw).Value
        Range("Q" & rw + 1 & ":" & "AB" & rw + 5).ClearContents
    End If
End If
End Sub

Sub insert1row()
If ActiveSheet.Name = "Data" Then
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        Dim rw As Integer
        rw = ActiveCell.Row
        Rows(rw + 1).Insert Shift:=xlDown
        Range("H" & rw + 1 & ":" & "J" & rw + 1).Value = Range("H" & rw & ":" & "J" & rw).Value
        Range("Q" & rw + 1).ClearContents
    End If
End If
End Sub

Sub insert1rowabove()
If ActiveSheet.Name = "Data" Then
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        Dim rw As Integer
        rw = ActiveCell.Row
        Rows(rw).Insert Shift:=xlDown
        Range("H" & rw & ":" & "J" & rw).Value = Range("H" & rw + 1 & ":" & "J" & rw + 1).Value
        Range("Q" & rw).ClearContents
    End If
End If

End Sub

Sub deleteactiverows()
    If ActiveSheet.Name = "Data" Then
        If ActiveCell.Row < 6 Or _
        ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
            MsgBox ("Please select a cell in the table first")
        Else
            Dim rng As Range
            Dim filt()
            Set rng = Selection.SpecialCells(xlCellTypeVisible)

            If Worksheets("Data").ListObjects("dataTable").AutoFilter.FilterMode = True Then
                Application.ScreenUpdating = False
                If InStr(1, rng.Address, ",") > 0 Then
                    rng.SpecialCells(xlCellTypeVisible).Delete
                    Exit Sub
                Else
                    MsgBox ("does not contain hidden rows")
                    MsgBox (rng.Address)
                    Call SaveListObjectFilters(Worksheets("Data").ListObjects("dataTable"), filt)
                    Rows(rng.Row & ":" & rng.Rows.Count + rng.Row - 1).Delete
                    Call RestoreListObjectFilters(Worksheets("Data").ListObjects("dataTable"), filt)
                End If
            Else
                Rows(rng.Row & ":" & rng.Rows.Count + rng.Row - 1).Delete
                Application.ScreenUpdating = True
            End If
        End If
    End If

End Sub

Sub reverseformat()
If ActiveSheet.Name = "Data" Then
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        rw = ActiveCell.Row
        If Range("M" & rw).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" Then
            Range("M" & rw).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            Range("Q" & rw & ":" & "AB" & rw).NumberFormat = "#,##0"
        ElseIf Range("M" & rw).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)" Then
            Range("M" & rw).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
            Range("Q" & rw & ":" & "AB" & rw).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        ElseIf Range("M" & rw).NumberFormat = "General" Then
            Range("M" & rw).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
            Range("Q" & rw & ":" & "AB" & rw).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        End If
    End If
End If
End Sub

Sub percentlineitem()
If ActiveSheet.Name = "Data" Then
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        Dim rw As Integer
        rw = ActiveCell.Row
        Range("M" & rw).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Range("Q" & rw & ":" & "AB" & rw).NumberFormat = "0%"
        
        If Range("prim_div_qty_Z1").Value <> "" Then
            Range("Q" & rw).Formula = "=prim_div_qty_Z1/prim_div_qty"
            Else: Range("Q" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z2").Value <> "" Then
            Range("R" & rw).Formula = "=prim_div_qty_Z2/prim_div_qty"
            Else: Range("R" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z3").Value <> "" Then
            Range("S" & rw).Formula = "=prim_div_qty_Z3/prim_div_qty"
            Else: Range("S" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z4").Value <> "" Then
            Range("T" & rw).Formula = "=prim_div_qty_Z4/prim_div_qty"
            Else: Range("T" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z5").Value <> "" Then
            Range("U" & rw).Formula = "=prim_div_qty_Z5/prim_div_qty"
            Else: Range("U" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z6").Value <> "" Then
            Range("V" & rw).Formula = "=prim_div_qty_Z6/prim_div_qty"
            Else: Range("V" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z7").Value <> "" Then
            Range("W" & rw).Formula = "=prim_div_qty_Z7/prim_div_qty"
            Else: Range("W" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z8").Value <> "" Then
            Range("X" & rw).Formula = "=prim_div_qty_Z8/prim_div_qty"
            Else: Range("X" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z9").Value <> "" Then
            Range("Y" & rw).Formula = "=prim_div_qty_Z9/prim_div_qty"
            Else: Range("Y" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z10").Value <> "" Then
            Range("Z" & rw).Formula = "=prim_div_qty_Z10/prim_div_qty"
            Else: Range("Z" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z11").Value <> "" Then
            Range("AA" & rw).Formula = "=prim_div_qty_Z11/prim_div_qty"
            Else: Range("AA" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z12").Value <> "" Then
            Range("AB" & rw).Formula = "=prim_div_qty_Z12/prim_div_qty"
            Else: Range("AB" & rw).Value = ""
        End If
    End If
End If
End Sub

Sub clearcomments()
    Cells.clearcomments
End Sub

Sub addbond()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        Dim rw As Integer
        rw = ActiveCell.Row
        Rows(rw + 1).Insert Shift:=xlDown
        rw = rw + 1
        Range("H" & rw).Value = "Z.70_Taxes_Permits_Insurance_and_Bonds"
        Range("I" & rw).Value = "Z.7070_Bond_Fees"
        Range("J" & rw).Value = Range("J" & rw - 1).Value
        Range("M" & rw).NumberFormat = "0.00%"
        Range("Q" & rw & ":" & "AB" & rw).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
        Range("M" & rw).Value = 0.01
        Range("L" & rw).Value = "Subcontractor Performance & Payment Bond"
     
        If Range("prim_div_qty_Z1").Value <> "" Then
            Range("Q" & rw).Formula = "=SUMIFS([ZONE1_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("Q" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z2").Value <> "" Then
            Range("R" & rw).Formula = "=SUMIFS([ZONE2_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("R" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z3").Value <> "" Then
            Range("S" & rw).Formula = "=SUMIFS([ZONE3_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("S" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z4").Value <> "" Then
            Range("T" & rw).Formula = "=SUMIFS([ZONE4_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("T" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z5").Value <> "" Then
            Range("U" & rw).Formula = "=SUMIFS([ZONE5_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("U" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z6").Value <> "" Then
            Range("V" & rw).Formula = "=SUMIFS([ZONE6_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("V" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z7").Value <> "" Then
            Range("W" & rw).Formula = "=SUMIFS([ZONE7_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("W" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z8").Value <> "" Then
            Range("X" & rw).Formula = "=SUMIFS([ZONE8_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("X" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z9").Value <> "" Then
            Range("Y" & rw).Formula = "=SUMIFS([ZONE9_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("Y" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z10").Value <> "" Then
            Range("Z" & rw).Formula = "=SUMIFS([ZONE10_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("Z" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z11").Value <> "" Then
            Range("AA" & rw).Formula = "=SUMIFS([ZONE11_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("AA" & rw).Value = ""
        End If
        If Range("prim_div_qty_Z12").Value <> "" Then
            Range("AB" & rw).Formula = "=SUMIFS([ZONE12_EXT],[CONTRACT ITEM],[@[CONTRACT ITEM]],[UNI  L3/L4],""<>Z.7070_Bond_Fees"")"
            Else: Range("AB" & rw).Value = ""
        End If
        
    End If
End Sub

Sub lineup()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        Application.ScreenUpdating = False
        Call SaveListObjectFilters(Worksheets("Data").ListObjects("dataTable"), filt)

        With Rows(ActiveCell.Row).EntireRow
            .Cut
            .Offset(-1).Insert
        End With
        ActiveCell.Offset(-1, 0).Select
        Call RestoreListObjectFilters(Worksheets("Data").ListObjects("dataTable"), filt)
        Application.ScreenUpdating = True
    End If
End Sub

Sub linedown()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        Application.ScreenUpdating = False
        Call SaveListObjectFilters(Worksheets("Data").ListObjects("dataTable"), filt)
        With Rows(ActiveCell.Row).EntireRow
            .Cut
            .Offset(2).Insert
        End With
        ActiveCell.Offset(1, 0).Select
        Call RestoreListObjectFilters(Worksheets("Data").ListObjects("dataTable"), filt)
        Application.ScreenUpdating = True
    End If
End Sub

Sub lineitemsummary()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        If Cells(ActiveCell.Row, 7).Value = "S" Then
            Cells(ActiveCell.Row, 7).Value = ""
        Else
            Cells(ActiveCell.Row, 7).Value = "S"
        End If
    End If
End Sub

Sub lineitemheading()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        If Cells(ActiveCell.Row, 7).Value = "H" Then
            Cells(ActiveCell.Row, 7).Value = ""
        Else
            Cells(ActiveCell.Row, 7).Value = "H"
        End If
    End If
End Sub

Sub lineiteminclude()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        If Cells(ActiveCell.Row, 7).Value = "%" Then
            Cells(ActiveCell.Row, 7).Value = ""
        Else
            Cells(ActiveCell.Row, 7).Value = "%"
        End If
    End If
End Sub

Sub lineitemexclude()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        If Cells(ActiveCell.Row, 7).Value = "X" Then
            Cells(ActiveCell.Row, 7).Value = ""
        Else
            Cells(ActiveCell.Row, 7).Value = "X"
        End If
    End If
End Sub

Sub lineitemnote()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
        MsgBox ("Please select a cell in the table first")
    Else
        If Cells(ActiveCell.Row, 7).Value = "*" Then
            Cells(ActiveCell.Row, 7).Value = ""
        Else
            Cells(ActiveCell.Row, 7).Value = "*"
        End If
    End If
End Sub

Sub permref()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Or _
    WorksheetFunction.IsFormula(ActiveCell) = False _
    Then
        MsgBox ("Please select a formula cell in the table first")
    Else
        Dim refrow As Integer
        Dim zone As Integer
        Dim zonecount As Integer
        Dim i As Integer
        Dim rw As Integer
        
        refrow = onlyDigits(ActiveCell.Formula)
        zone = 1
        zonecount = WorksheetFunction.CountA(Range("zonecountrange"))
        rw = ActiveCell.Row
        
        For i = 17 To 16 + zonecount Step 1
            Cells(rw, i).Formula = "=VLOOKUP(" & """" & Cells(refrow, 5).Value & """" & ",dataTable[[GUID]:[ZONE12]]," & zone + 12 & ",FALSE)"
            zone = zone + 1
        Next i

    End If
End Sub

Function onlyDigits(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    onlyDigits = retval
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Sub:      SaveListObjectFilters
' Purpose:  Save filter on worksheet
' Returns:  wks.AutoFilterMode when function entered
' Source: http://stackoverflow.com/questions/9489126/in-excel-vba-how-do-i-save-        restore-a-user-defined-filter
'
' Arguments:
'   [Name]      [Type]  [Description]
'   wks         I/P     Worksheet that filter may reside on
'   FilterRange O/P     Range on which filter is applied as string; "" if no filter
'   FilterCache O/P     Variant dynamic array in which to save filter
'
' Author:   Based on MS Excel AutoFilter Object help file
'
' Modifications:
' 2006/12/11 Phil Spencer: Adapted as general purpose routine
' 2007/03/23 PJS: Now turns off .AutoFilterMode
' 2013/03/13 PJS: Initial mods for XL14, which has more operators
' 2013/05/31 P.H.: Changed to save list-object filters

Public Function SaveListObjectFilters(lo As ListObject, FilterCache()) As Boolean

Dim ii As Long

filterRange = ""
    With lo.AutoFilter
        filterRange = .Range.Address
        With .Filters
            ReDim FilterCache(1 To .Count, 1 To 3)
            For ii = 1 To .Count
                With .Item(ii)
                    If .On Then
#If False Then ' XL11 code
                        FilterCache(ii, 1) = .Criteria1
                        If .Operator Then
                            FilterCache(ii, 2) = .Operator
                            FilterCache(ii, 3) = .Criteria2
                        End If
#Else   ' first pass XL14
                        Select Case .Operator

                        Case 1, 2   'xlAnd, xlOr
                            FilterCache(ii, 1) = .Criteria1
                            FilterCache(ii, 2) = .Operator
                            FilterCache(ii, 3) = .Criteria2

                        Case 0, 3 To 7 ' no operator, xlTop10Items, _
xlBottom10Items, xlTop10Percent, xlBottom10Percent, xlFilterValues
                            FilterCache(ii, 1) = .Criteria1
                            FilterCache(ii, 2) = .Operator

                        Case Else    ' These are not correctly restored; there's someting in Criteria1 but can't save it.
                            FilterCache(ii, 2) = .Operator
                            ' FilterCache(ii, 1) = .Criteria1   ' <-- Generates an error
                            ' No error in next statement, but couldn't do restore operation
                            ' Set FilterCache(ii, 1) = .Criteria1

                        End Select
#End If
                    End If
                End With ' .Item(ii)
            Next
        End With ' .Filters
    lo.AutoFilter.ShowAllData
    End With ' wks.AutoFilter
End Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Sub:      RestoreListObjectFilters
' Purpose:  Restore filter on listobject
' Source: http://stackoverflow.com/questions/9489126/in-excel-vba-how-do-i-save-restore-a-user-defined-filter
' Arguments:
'   [Name]      [Type]  [Description]
'   wks         I/P     Worksheet that filter resides on
'   FilterRange I/P     Range on which filter is applied
'   FilterCache I/P     Variant dynamic array containing saved filter
'
' Author:   Based on MS Excel AutoFilter Object help file
'
' Modifications:
' 2006/12/11 Phil Spencer: Adapted as general purpose routine
' 2013/03/13 PJS: Initial mods for XL14, which has more operators
' 2013/05/31 P.H.: Changed to restore list-object filters
'
' Comments:
'----------------------------
Public Sub RestoreListObjectFilters(lo As ListObject, FilterCache())
Dim col As Long

If lo.Range.Address <> "" Then
    For col = 1 To UBound(FilterCache(), 1)

#If False Then  ' XL11
        If Not IsEmpty(FilterCache(col, 1)) Then
            If FilterCache(col, 2) Then
                lo.AutoFilter field:=col, _
                    Criteria1:=FilterCache(col, 1), _
                        Operator:=FilterCache(col, 2), _
                    Criteria2:=FilterCache(col, 3)
            Else
                lo.AutoFilter field:=col, _
                    Criteria1:=FilterCache(col, 1)
            End If
        End If
#Else

        If Not IsEmpty(FilterCache(col, 2)) Then
            Select Case FilterCache(col, 2)

            Case 0  ' no operator
                lo.Range.AutoFilter field:=col, _
                    Criteria1:=FilterCache(col, 1) ' Do NOT reload 'Operator'

            Case 1, 2   'xlAnd, xlOr
                lo.Range.AutoFilter field:=col, _
                    Criteria1:=FilterCache(col, 1), _
                    Operator:=FilterCache(col, 2), _
                    Criteria2:=FilterCache(col, 3)

            Case 3 To 6 ' xlTop10Items, xlBottom10Items, xlTop10Percent,     xlBottom10Percent
#If True Then
                lo.Range.AutoFilter field:=col, _
                    Criteria1:=FilterCache(col, 1) ' Do NOT reload 'Operator' , it doesn't work
                ' wks.AutoFilter.Filters.Item(col).Operator = FilterCache(col, 2)
#Else ' Trying to restore Operator as well as Criteria ..
                ' Including the 'Operator:=' arguement leads to error.
                ' Criteria1 is expressed as if for a FALSE .Operator
                lo.Range.AutoFilter field:=col, _
                    Criteria1:=FilterCache(col, 1), _
                    Operator:=FilterCache(col, 2)
#End If

            Case 7  'xlFilterValues
                lo.Range.AutoFilter field:=col, _
                    Criteria1:=FilterCache(col, 1), _
                    Operator:=FilterCache(col, 2)

#If False Then ' Switch on filters on cell formats
' These statements restore the filter, but cannot reset the pass Criteria, so the filter hides all data.
' Leave it off instead.
            Case Else   ' (Various filters on data format)
                lo.RangeAutoFilter field:=col, _
                    Operator:=FilterCache(col, 2)
#End If ' Switch on filters on cell formats

            End Select
        End If

#End If     ' XL11 / XL14
    Next col
End If
End Sub

