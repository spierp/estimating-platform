Attribute VB_Name = "detailDHeadings"
Sub createHeadings()
'pb progress 35%
pb.Repaint

'CREATE HEADINGS
pb.AddCaption "Creating Level 1 Headings... (this might take a minute)"
Dim x As Long, lastrow As Long
lastrow = Cells(Rows.count, 16).End(xlUp).Row

For x = lastrow To 7 Step -1
    If Cells(x, 1).Font.Bold = False _
    And Cells(x, 1).Value <> Cells(x - 1, 1).Value _
    And Cells(x - 1, 1).Value <> "" _
    And Cells(x, 1).Value <> "" Then
        Rows(x).EntireRow.Insert
        lastrow = lastrow + 1
        Rows(x).RowHeight = 22
        Rows(x).Font.Bold = True
        Cells(x, 1).Value = Cells(x + 1, 1).Value
    End If

Next x
pb.AddProgress 5

pb.AddCaption "Creating Level 2 Headings... (this might take a longer minute)"
For x = lastrow To 7 Step -1
    If Cells(x, 2).Font.Bold = False _
    And Cells(x, 2).Value <> Cells(x - 1, 2).Value _
    And Cells(x, 2).Value <> Cells(x - 2, 2).Value _
    And Cells(x, 2).Value <> "" Then
        Rows(x).EntireRow.Insert
        lastrow = lastrow + 1
        Rows(x).RowHeight = 18
        Rows(x).Font.Bold = True
        Cells(x, 2).Value = Cells(x + 1, 2).Value
    End If
Next x
pb.AddProgress 10

pb.AddCaption "Creating Level 3 Headings... (this is the longest minute)"
For x = lastrow To 7 Step -1
    If Cells(x, 3).Font.Bold = False _
    And Cells(x, 3).Value <> Cells(x - 1, 3).Value _
    And Cells(x, 3).Value <> "" Then
        Rows(x).EntireRow.Insert
        lastrow = lastrow + 1
        Rows(x).RowHeight = 14
        Rows(x).Font.Bold = True
        Cells(x, 3).Value = Cells(x + 1, 3).Value
    End If
    
Next x
pb.AddProgress 30

'DELETE EXTRA CODES
pb.AddCaption "Deleting extra code references..."

Dim dat As Variant
Dim rng As range
Dim i As Long

Set rng = range("$A$7", Cells(Rows.count, "C").End(xlUp)).Cells
dat = rng.Value
For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 1) <> "" And dat(i, 2) <> "" And dat(i, 3) <> "" Then
    dat(i, 1) = ""
    dat(i, 2) = ""
    dat(i, 3) = ""
    End If
Next
rng.Value = dat

Erase dat
pb.AddProgress 5

'ADD ROW NUMBERS
pb.AddCaption "Adding row numbers..."

Dim lineCount As Integer
lineCount = 0

Set rng = range("$A$7", Cells(Rows.count, "L").End(xlUp)).Cells
dat = rng.Value
For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 12) <> "" Then
    lineCount = lineCount + 1
    dat(i, 1) = lineCount
    End If
Next
rng.Value = dat
pb.AddProgress 5

End Sub
