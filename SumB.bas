Attribute VB_Name = "SumB"
Sub sumColumns()
pb.AddCaption "Hiding markup rows that are not applicable..."
    
    Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False
    
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
        range("G7").EntireColumn.Hidden = True
    End If
    
    If ThisWorkbook.Names("sum_show_prim_div").RefersToRange(1, 1).Value = "No" Then
        range("E7").EntireColumn.Hidden = True
        range("J7").EntireColumn.Hidden = True
        range("N7").EntireColumn.Hidden = True
        range("R7").EntireColumn.Hidden = True
        range("V7").EntireColumn.Hidden = True
        range("Z7").EntireColumn.Hidden = True
        range("AD7").EntireColumn.Hidden = True
        range("AH7").EntireColumn.Hidden = True
        range("AL7").EntireColumn.Hidden = True
        range("AP7").EntireColumn.Hidden = True
        range("AT7").EntireColumn.Hidden = True
        range("AX7").EntireColumn.Hidden = True
        range("BB7").EntireColumn.Hidden = True
    End If
    If ThisWorkbook.Names("sum_show_sec_div").RefersToRange(1, 1).Value = "No" Then
        range("F7").EntireColumn.Hidden = True
        range("K7").EntireColumn.Hidden = True
        range("O7").EntireColumn.Hidden = True
        range("S7").EntireColumn.Hidden = True
        range("W7").EntireColumn.Hidden = True
        range("AA7").EntireColumn.Hidden = True
        range("AE7").EntireColumn.Hidden = True
        range("AI7").EntireColumn.Hidden = True
        range("AM7").EntireColumn.Hidden = True
        range("AQ7").EntireColumn.Hidden = True
        range("AU7").EntireColumn.Hidden = True
        range("AY7").EntireColumn.Hidden = True
        range("BC7").EntireColumn.Hidden = True
    End If
    
    If ThisWorkbook.Names("name_Z2").RefersToRange(1, 1).Value = "" Then
        range("H10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z3").RefersToRange(1, 1).Value = "" Then
        range("P10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z4").RefersToRange(1, 1).Value = "" Then
        range("T10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z5").RefersToRange(1, 1).Value = "" Then
        range("X10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z6").RefersToRange(1, 1).Value = "" Then
        range("AB10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z7").RefersToRange(1, 1).Value = "" Then
        range("AF10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z8").RefersToRange(1, 1).Value = "" Then
        range("AJ10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z9").RefersToRange(1, 1).Value = "" Then
        range("AN10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z10").RefersToRange(1, 1).Value = "" Then
        range("AR10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z11").RefersToRange(1, 1).Value = "" Then
        range("AV10:BB10").EntireColumn.Hidden = True
    ElseIf ThisWorkbook.Names("name_Z12").RefersToRange(1, 1).Value = "" Then
        range("AZ10:BB10").EntireColumn.Hidden = True
    End If
    
If ThisWorkbook.Names("name_Z2").RefersToRange(1, 1).Value = "" Then
    range("$BB$4:$BB$6").HorizontalAlignment = xlRight
    Rows("7:8").Hidden = True
End If
    
If ActiveSheet.Name = "brkSum" Or ActiveSheet.Name = "altSum" Then
    Rows("7:8").Hidden = True
End If
    
pb.AddProgress 35
    
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
        .FooterMargin = Application.InchesToPoints(0.15)
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
    
    range("D7").Select

pb.AddProgress 5
    
End Sub


