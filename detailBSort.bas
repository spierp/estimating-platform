Attribute VB_Name = "detailBSort"
Sub sortAndFormat()
pb.Repaint
'pb 15% complete
pb.AddCaption "Sorting table columns..."
   
'COLUMN RELOCATION
    Columns("A:G").Cut
    Columns("K:K").Insert Shift:=xlToRight
    Rows(7).Delete Shift:=xlUp
   
'FORMAT COLUMN WIDTHS
    Columns("A:B").ColumnWidth = 5
    Columns("C:C").ColumnWidth = 1
    Columns("L:L").ColumnWidth = 55
    Columns("M:M").ColumnWidth = 13
    Columns("N:N").ColumnWidth = 6
    Columns("P:P").ColumnWidth = 16
    Columns("Q:AB").ColumnWidth = 11
    Columns("AC:AN").ColumnWidth = 13


'HIDE COLUMNS
    Columns("D:K").EntireColumn.Hidden = True
    If WorksheetFunction.CountA(Range("B7:B10")) < 2 Then
        Columns("B:B").EntireColumn.Hidden = True
    End If
    If WorksheetFunction.CountA(Range("C7:C10")) < 2 Then
        Columns("C:C").EntireColumn.Hidden = True
    End If
   
'FORMAT FONT & ALIGNMENT
    With Columns("O:AZ")
        .HorizontalAlignment = xlRight
    End With
    
    With Rows("6:6")
        .Font.Bold = True
        .RowHeight = 22
        .HorizontalAlignment = xlCenter
    End With

    With Columns("P:P")
            .Font.Bold = True
    End With

pb.AddProgress 10
'pb 20% complete

'DELETE EXTRA ROWS
Dim x As Long, lastrow As Long
lastrow = Cells(Rows.Count, 16).End(xlUp).Row
For x = lastrow To 1 Step -1
    If Cells(x, 1).Value = "" And Cells(x, 2) = "" And Cells(x, 3) = "" And Cells(x, 16) <> "" Then
        Rows(x).Delete
    End If
Next x
      
End Sub


