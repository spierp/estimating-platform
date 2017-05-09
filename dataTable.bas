Attribute VB_Name = "dataTable"
Sub dataSort()
Attribute dataSort.VB_ProcData.VB_Invoke_Func = " \n14"
    range("dataTable").Select
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=range("dataTable[CONTRACT ITEM]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=range("dataTable[UNI L2]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort.SortFields.Add _
        Key:=range("dataTable[UNI  L3/L4]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Data").ListObjects("dataTable").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    range("A1").Select
End Sub

Sub insert5rows()
If ActiveCell.Row < 6 Or _
ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).range.Rows.count Then
    MsgBox ("Please select a cell in the table first")
Else
    Dim rw As Integer
    rw = ActiveCell.Row
    Rows(rw + 1 & ":" & rw + 5).Insert Shift:=xlDown
    range("H" & rw + 1 & ":" & "J" & rw + 5).Value = range("H" & rw & ":" & "J" & rw).Value
    range("Q" & rw + 1 & ":" & "AB" & rw + 5).ClearContents
End If
End Sub

Sub insert1row()
If ActiveCell.Row < 6 Or _
ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).range.Rows.count Then
    MsgBox ("Please select a cell in the table first")
Else
    Dim rw As Integer
    rw = ActiveCell.Row
    Rows(rw + 1).Insert Shift:=xlDown
    range("H" & rw + 1 & ":" & "J" & rw + 1).Value = range("H" & rw & ":" & "J" & rw).Value
    range("Q" & rw + 1).ClearContents
End If
End Sub

Sub reverseFormat()
If ActiveCell.Row < 6 Or _
ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).range.Rows.count Then
    MsgBox ("Please select a cell in the table first")
Else
    rw = ActiveCell.Row
    If range("M" & rw).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" Then
        range("M" & rw).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        range("Q" & rw & ":" & "AB" & rw).NumberFormat = "#,##0"
    ElseIf range("M" & rw).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)" Then
        range("M" & rw).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
        range("Q" & rw & ":" & "AB" & rw).NumberFormat = "_($* #,##0_);_($* (#,##0);_($* ""-""??_);_(@_)"
    End If
End If
End Sub

Sub percentlineitem()
    If ActiveCell.Row < 6 Or _
    ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).range.Rows.count Then
        MsgBox ("Please select a cell in the table first")
    Else
        Dim rw As Integer
        rw = ActiveCell.Row
        range("M" & rw).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        range("Q" & rw & ":" & "AB" & rw).NumberFormat = "0%"
        
        If range("prim_div_qty_Z1").Value <> "" Then
            range("Q" & rw).Formula = "=prim_div_qty_Z1/prim_div_qty"
            Else: range("Q" & rw).Value = ""
        End If
        If range("prim_div_qty_Z2").Value <> "" Then
            range("R" & rw).Formula = "=prim_div_qty_Z2/prim_div_qty"
            Else: range("R" & rw).Value = ""
        End If
        If range("prim_div_qty_Z3").Value <> "" Then
            range("S" & rw).Formula = "=prim_div_qty_Z3/prim_div_qty"
            Else: range("S" & rw).Value = ""
        End If
        If range("prim_div_qty_Z4").Value <> "" Then
            range("T" & rw).Formula = "=prim_div_qty_Z4/prim_div_qty"
            Else: range("T" & rw).Value = ""
        End If
        If range("prim_div_qty_Z5").Value <> "" Then
            range("U" & rw).Formula = "=prim_div_qty_Z5/prim_div_qty"
            Else: range("U" & rw).Value = ""
        End If
        If range("prim_div_qty_Z6").Value <> "" Then
            range("V" & rw).Formula = "=prim_div_qty_Z6/prim_div_qty"
            Else: range("V" & rw).Value = ""
        End If
        If range("prim_div_qty_Z7").Value <> "" Then
            range("W" & rw).Formula = "=prim_div_qty_Z7/prim_div_qty"
            Else: range("W" & rw).Value = ""
        End If
        If range("prim_div_qty_Z8").Value <> "" Then
            range("X" & rw).Formula = "=prim_div_qty_Z8/prim_div_qty"
            Else: range("X" & rw).Value = ""
        End If
        If range("prim_div_qty_Z9").Value <> "" Then
            range("Y" & rw).Formula = "=prim_div_qty_Z9/prim_div_qty"
            Else: range("Y" & rw).Value = ""
        End If
        If range("prim_div_qty_Z10").Value <> "" Then
            range("Z" & rw).Formula = "=prim_div_qty_Z10/prim_div_qty"
            Else: range("Z" & rw).Value = ""
        End If
        If range("prim_div_qty_Z11").Value <> "" Then
            range("AA" & rw).Formula = "=prim_div_qty_Z11/prim_div_qty"
            Else: range("AA" & rw).Value = ""
        End If
        If range("prim_div_qty_Z12").Value <> "" Then
            range("AB" & rw).Formula = "=prim_div_qty_Z12/prim_div_qty"
            Else: range("AB" & rw).Value = ""
        End If
    End If
End Sub
