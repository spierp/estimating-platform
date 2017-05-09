Attribute VB_Name = "detailFPageFormat"
Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Sub pageFormat()
pb.Repaint

'Add Header Info
pb.AddCaption "Creating Page Header..."
'Insert Logo
Worksheets("dashboard").Shapes("full_logo").Copy
range("A1").Select
ActiveSheet.Paste
Selection.ShapeRange.ScaleHeight 0.5715085068, msoFalse, msoScaleFromTopLeft

rightColumn = Col_Letter(WorksheetFunction.CountA(range("A6:AZ6")))

'Center Header
range("C1").Value = UCase(ThisWorkbook.Names("project_name").RefersToRange(1, 1).Value)
range("C2").Value = UCase(ThisWorkbook.Names("client_name").RefersToRange(1, 1).Value)
range("C3").Value = UCase(ThisWorkbook.Names("estimate_name").RefersToRange(1, 1).Value)

range("C1:" & rightColumn & "4").Select

    With Selection
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
    End With
    With Selection.Font
        .Bold = True
    End With

range("C2").Font.Underline = True

'Right Header

If ActiveSheet.Name = "altDetail" Then
    sheetName = "ALTERNATES DETAIL"
ElseIf ActiveSheet.Name = "brkDetail" Then
    sheetName = "BREAK-OUT DETAIL"
ElseIf ActiveSheet.Name = "subDetail" Then
    sheetName = "SUBCONTRACTOR DETAIL"
ElseIf ActiveSheet.Name = "tradeDetail" Then
    sheetName = "LINE ITEM DETAIL - SORTED BY TRADE"
ElseIf ActiveSheet.Name = "uniDetail" Then
    sheetName = "LINE ITEM DETAIL - SORTED BY SYSTEM"
End If

range(rightColumn & "1:" & rightColumn & "3").Select
    With Selection
        .HorizontalAlignment = xlRight
    End With

range(rightColumn & "1").Value = sheetName
range(rightColumn & "2").Value = ThisWorkbook.Names("estimate_date").RefersToRange(1, 1).Value
range(rightColumn & "2").NumberFormat = "dd/mm/yyyy"
range(rightColumn & "3").Value = 1
range(rightColumn & "3").Font.Color = RGB(255, 255, 255)

Rows("1").Select
    With Selection.Font
        .Size = 11
    End With
    
Rows("7").Select
ActiveWindow.FreezePanes = True
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
Selection.RowHeight = 12

range("A8:C10").HorizontalAlignment = xlLeft

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

