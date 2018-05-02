Attribute VB_Name = "SumB"
Sub sumColumns()

Dim rng As Range
Dim breakcount As Integer
Dim altcount As Integer
Dim vArray As Variant
Dim lastrow As Integer

lastrow = Worksheets("Data").ListObjects("dataTable").TotalsRowRange.Row

pb.AddCaption "Hiding markup rows that are not applicable..."
    
Cells.Select
Selection.EntireColumn.Hidden = False
Selection.EntireRow.Hidden = False

'HIDE IRRELEVANT MARK-UPS
ActiveSheet.Columns(3).Find("COST OF WORK - SUBTOTAL").Select
ActiveCell.Offset(2, 0).Select

Do Until ActiveCell.Value = ""
    If ActiveCell.Offset(0, 1) = "0" Then
        ActiveCell.Rows("1:1").EntireRow.Select
        Selection.EntireRow.Hidden = True
        ActiveCell.Offset(0, 2).Select
    End If
    ActiveCell.Offset(1, 0).Select
Loop

pb.AddProgress 10
pb.AddCaption "Hiding columns that are not applicable..."

If ThisWorkbook.Names("sum_show_comments").RefersToRange(1, 1).Value = "No" Then
    Range("G7").EntireColumn.Hidden = True
End If

If ThisWorkbook.Names("sum_show_prim_div").RefersToRange(1, 1).Value = "No" Then
    Range("E7").EntireColumn.Hidden = True
    Range("J7").EntireColumn.Hidden = True
    Range("N7").EntireColumn.Hidden = True
    Range("R7").EntireColumn.Hidden = True
    Range("V7").EntireColumn.Hidden = True
    Range("Z7").EntireColumn.Hidden = True
    Range("AD7").EntireColumn.Hidden = True
    Range("AH7").EntireColumn.Hidden = True
    Range("AL7").EntireColumn.Hidden = True
    Range("AP7").EntireColumn.Hidden = True
    Range("AT7").EntireColumn.Hidden = True
    Range("AX7").EntireColumn.Hidden = True
    Range("BB7").EntireColumn.Hidden = True
End If
If ThisWorkbook.Names("sum_show_sec_div").RefersToRange(1, 1).Value = "No" Then
    Range("F7").EntireColumn.Hidden = True
    Range("K7").EntireColumn.Hidden = True
    Range("O7").EntireColumn.Hidden = True
    Range("S7").EntireColumn.Hidden = True
    Range("W7").EntireColumn.Hidden = True
    Range("AA7").EntireColumn.Hidden = True
    Range("AE7").EntireColumn.Hidden = True
    Range("AI7").EntireColumn.Hidden = True
    Range("AM7").EntireColumn.Hidden = True
    Range("AQ7").EntireColumn.Hidden = True
    Range("AU7").EntireColumn.Hidden = True
    Range("AY7").EntireColumn.Hidden = True
    Range("BC7").EntireColumn.Hidden = True
End If
       
If ActiveSheet.Name = "brkSum" Then
    breakcount = Sheets("Data").Range("dataTable[[#Totals],[BRK]]").Value
    Set rng = Worksheets("Data").Range("C5:C" & lastrow)
    pb.AddCaption (breakcount & " break-outs found...")
    If breakcount < 2 Then
        Range("L10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 3 Then
        Range("P10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 4 Then
        Range("T10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 5 Then
        Range("X10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 6 Then
        Range("AB10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 7 Then
        Range("AF10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 8 Then
        Range("AJ10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 9 Then
        Range("AN10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 10 Then
        Range("AR10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 11 Then
        Range("AV10:BB10").EntireColumn.Hidden = True
    ElseIf breakcount < 12 Then
        Range("AZ10:BB10").EntireColumn.Hidden = True
    End If
End If
        
If ActiveSheet.Name = "altSum" Then
    altcount = Sheets("Data").Range("dataTable[[#Totals],[ALT]]").Value
    Set rng = Worksheets("Data").Range("D5:D" & lastrow)
    pb.AddCaption (altcount & " alternates found...")
    If altcount < 2 Then
        Range("L10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 3 Then
        Range("P10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 4 Then
        Range("T10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 5 Then
        Range("X10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 6 Then
        Range("AB10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 7 Then
        Range("AF10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 8 Then
        Range("AJ10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 9 Then
        Range("AN10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 10 Then
        Range("AR10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 11 Then
        Range("AV10:BB10").EntireColumn.Hidden = True
    ElseIf altcount < 12 Then
        Range("AZ10:BB10").EntireColumn.Hidden = True
    End If
End If
        
pb.AddProgress 10
        
If ActiveSheet.Name = "brkSum" Or ActiveSheet.Name = "altSum" Then

    pb.AddCaption "Adding tags..."
    Rows("7:9").Hidden = True
    Range("D10:G10").EntireColumn.Hidden = True

    Dim baArray As Variant
    Dim i As Integer
    Dim x As Integer
    
    rng.AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Worksheets("clipboard").Range("CA1"), Unique:=True
    
    baArray = Sheets("clipboard").Range("CA2:CA15").Value
    
    Sheets("clipboard").Range("CA1:CA15").ClearContents
    
    x = 9

    For i = LBound(baArray, 1) To UBound(baArray, 1)
        If baArray(i, 1) <> "" Then
            ActiveSheet.Cells(10, x) = baArray(i, 1)
            x = x + 4
        End If
    Next
    
    Erase baArray
        
    Worksheets("Data").ListObjects("dataTable").ShowAutoFilter = True
    pb.AddProgress 30

End If
        
If ActiveSheet.Name = "tradeSum" Or ActiveSheet.Name = "uni2Sum" Or ActiveSheet.Name = "uni34Sum" Then
        
    If ThisWorkbook.Names("name_Z2").RefersToRange(1, 1).Value = "" Then
        Range("H10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z3").RefersToRange(1, 1).Value = "" Then
        Range("P10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z4").RefersToRange(1, 1).Value = "" Then
        Range("T10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z5").RefersToRange(1, 1).Value = "" Then
        Range("X10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z6").RefersToRange(1, 1).Value = "" Then
        Range("AB10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z7").RefersToRange(1, 1).Value = "" Then
        Range("AF10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z8").RefersToRange(1, 1).Value = "" Then
        Range("AJ10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z9").RefersToRange(1, 1).Value = "" Then
        Range("AN10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z10").RefersToRange(1, 1).Value = "" Then
        Range("AR10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z11").RefersToRange(1, 1).Value = "" Then
        Range("AV10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z12").RefersToRange(1, 1).Value = "" Then
        Range("AZ10:BB10").EntireColumn.Hidden = True
    End If
    
    If ThisWorkbook.Names("name_Z2").RefersToRange(1, 1).Value = "" Then
        'range("$BB$4:$BB$6").HorizontalAlignment = xlRight
        Rows("7:8").Hidden = True
    End If
pb.AddProgress 5
End If

pb.AddProgress 30

Sheets("clipboard").Visible = False
End Sub

Sub sumPageSetup()

pb.AddCaption "Configuring Print Setup..."
Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$11"
        .PrintTitleColumns = ""
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "Page &P of &N"
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.17)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = False
        If ThisWorkbook.Names("page_orientation").RefersToRange(1, 1).Value = "Portrait" Then
            .Orientation = xlPortrait
        Else
            .Orientation = xlLandscape
        End If
        .Draft = False
        If ThisWorkbook.Names("page_size").RefersToRange(1, 1).Value = "Letter" Then
            .PaperSize = xlPaperLetter
        ElseIf ThisWorkbook.Names("page_size").RefersToRange(1, 1).Value = "Legal" Then
            .PaperSize = xlPaperLegal
        Else
            .PaperSize = xlPaperTabloid
        End If
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 5
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.text = ""
        .EvenPage.CenterHeader.text = ""
        .EvenPage.RightHeader.text = ""
        .EvenPage.LeftFooter.text = ""
        .EvenPage.CenterFooter.text = ""
        .EvenPage.RightFooter.text = ""
        .FirstPage.LeftHeader.text = ""
        .FirstPage.CenterHeader.text = ""
        .FirstPage.RightHeader.text = ""
        .FirstPage.LeftFooter.text = ""
        .FirstPage.CenterFooter.text = ""
        .FirstPage.RightFooter.text = ""
    End With
    Application.PrintCommunication = True
    
    Range("D7").Select

pb.AddProgress 5
    
End Sub


