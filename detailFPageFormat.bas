Attribute VB_Name = "detailFPageFormat"
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Sub pageFormat()
pb.Repaint

'ADD HEADER INFO
pb.AddCaption "Creating Page Header..."

'INSERT LOGO
Worksheets("dashboard").Shapes("full_logo").Copy
Range("A1").Select
ActiveSheet.Paste
Selection.ShapeRange.ScaleHeight 0.5715085068, msoFalse, msoScaleFromTopLeft

If WorksheetFunction.CountA(Range("A6:AZ6")) < 19 Then
    rightColumn = Col_Letter(16)
    Columns("Q:R").EntireColumn.Hidden = True
Else
    rightColumn = Col_Letter(WorksheetFunction.CountA(Range("A6:AZ6")))
End If

'CENTER HEADER
Range("C1").Value = UCase(ThisWorkbook.Names("project_name").RefersToRange(1, 1).Value)
Range("C2").Value = UCase(ThisWorkbook.Names("client_name").RefersToRange(1, 1).Value)
Range("C3").Value = UCase(ThisWorkbook.Names("estimate_name").RefersToRange(1, 1).Value)

Range("C1:" & rightColumn & "4").Select

    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
    With Selection.Font
        .Bold = True
    End With

Range("C2").Font.Underline = True

'RIGHT HEADER
If ActiveSheet.Name = "altDetail" Then
    sheetname = "ALTERNATES DETAIL"
ElseIf ActiveSheet.Name = "brkDetail" Then
    sheetname = "BREAK-OUT DETAIL"
ElseIf ActiveSheet.Name = "subDetail" Then
    sheetname = "SUBCONTRACTOR DETAIL"
ElseIf ActiveSheet.Name = "tradeDetail" Then
    sheetname = "LINE ITEM DETAIL - SORTED BY TRADE"
ElseIf ActiveSheet.Name = "uniDetail" Then
    sheetname = "LINE ITEM DETAIL - SORTED BY SYSTEM"
End If

Range(rightColumn & "1:" & rightColumn & "3").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With

Range(rightColumn & "1").Value = sheetname
Range(rightColumn & "2").Value = ThisWorkbook.Names("estimate_date").RefersToRange(1, 1).Value
Range(rightColumn & "2").NumberFormat = "mm/dd/yyyy"
Range(rightColumn & "3").Value = 1
Range(rightColumn & "3").Font.Color = RGB(255, 255, 255)

Rows("1").Select
    With Selection.Font
        .Size = 11
    End With
    
Rows("7").Select
ActiveWindow.FreezePanes = True
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.RowHeight = 12

Range("A8:C10").HorizontalAlignment = xlLeft
Range("A9:A10").HorizontalAlignment = xlRight
Columns("A:B").ColumnWidth = 5
Columns("C").ColumnWidth = 1

pb.AddProgress 3

'CONFIGURING PRINT SETUP
pb.AddCaption "Configuring Print Setup..."

Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = "$1:$7"
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
        If ActiveSheet.Name = "brkDetail" Or ActiveSheet.Name = "altDetail" Then
            .FitToPagesTall = False
        Else
            .FitToPagesTall = 100
        End If
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
   
End Sub

