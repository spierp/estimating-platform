VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} qtouserform 
   Caption         =   "Import Quantities"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11520
   OleObjectBlob   =   "qtouserform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "qtouserform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim progress As Double, maxProgress As Double, maxWidth As Long, startTime As Double

Dim lastrow As Integer, datarowmarker As Integer, qtorowmarker As Integer, i As Integer, max As Long, zonecount As Integer
Dim matchcount As Integer, dat As Variant, rng As Range, qtowb As String, completed As Integer, newcount As Integer
Dim dataloopclosed As Boolean

Dim varTemp As Variant


Public Sub Initialize(qtowb As String)
'Initialize and shor progress bar

Call dynamo(qtowb)
If fatalerror = True Then
    Application.Workbooks(wbThis).Worksheets("Dashboard").Activate
    Range("A1").Select
    Me.Hide
    Exit Sub
End If

Dim max As Long
Dim rw As Integer

Application.Workbooks(qtowb).Worksheets("QTO").Activate

zonecount = WorksheetFunction.CountA(Range("G1:R1"))
max = Cells(Rows.Count, 4).End(xlUp).Row - 1

maxProgress = max:  maxWidth = lBar.Width:    lBar.Width = 0
lProgress.Caption = "0"
lcount.Caption = max
lzones.Caption = zonecount
lineitemframe.Caption = "Existing Line Item"

Range("A1").CurrentRegion.Select
Set rng = Selection.Cells
dat = rng.Value


Application.Workbooks(wbThis).Worksheets("Data").Activate
Cells.clearcomments

lastrow = Cells(Rows.Count, 1).End(xlUp).Row - 1
datarowmarker = 6

Me.Show False

Call dataloop

End Sub
Public Sub AddProgress(Optional inc As Long = 1)
'Increase progress by an increment
    completed = completed + 1
    progress = progress + inc
    If progress > maxProgress Then progress = maxProgress
    lBar.Width = CLng(CDbl(progress) / maxProgress * maxWidth)
    lProgress.Caption = "" & CLng(CDbl(progress) / maxProgress * 100) & "%"
    lcompleted.Caption = completed
    If progress = maxProgress Then Me.Hide
    DoEvents
End Sub

Private Sub CommandButtonImportFlag_Click()

If zonecount = 1 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7)))
ElseIf zonecount = 2 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8)))
ElseIf zonecount = 3 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9)))
ElseIf zonecount = 4 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10)))
ElseIf zonecount = 5 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10, 11)))
ElseIf zonecount = 6 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10, 11, 12)))
ElseIf zonecount = 7 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13)))
ElseIf zonecount = 8 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14)))
ElseIf zonecount = 9 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15)))
ElseIf zonecount = 10 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16)))
ElseIf zonecount = 11 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17)))
ElseIf zonecount = 12 Then
    varTemp = Application.Index(dat, qtorowmarker, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18)))
End If

If dataloopclosed = False Then
    Cells(datarowmarker, 14).Value = dat(qtorowmarker, 5)
    
    Cells(datarowmarker, 15).AddComment _
    "Previous QTO = " & Format(Cells(datarowmarker, 15).Value, "###,##0") & " " & Cells(datarowmarker, 14).Value
    
    Range("Q" & datarowmarker).Resize(LBound(varTemp), UBound(varTemp)) = Application.Transpose(varTemp)
    
Else
    
    rw = Range("L6").End(xlDown).Row + 1
    Sheets("Data").Range("A" & rw).EntireRow.Insert
    Sheets("Data").Range("Q" & rw).Resize(LBound(varTemp), UBound(varTemp)) = Application.Transpose(varTemp)
    
    Cells(rw, 14).Value = dat(qtorowmarker, 5)
    Cells(rw, 12).Value = dat(qtorowmarker, 4)
    Cells(rw, 10).Value = dat(qtorowmarker, 3)
    Cells(rw, 9).Value = dat(qtorowmarker, 2)

    dat(qtorowmarker, 1) = "imported"

End If

dat(qtorowmarker, 1) = "imported & flagged"

Call finishclick

End Sub

Private Sub CommandButtonSkip_Click()

dat(qtorowmarker, 1) = "skipped"
Call finishclick

End Sub

Private Sub CommandButtonSkipFlag_Click()

If dataloopclosed = False Then
    Cells(datarowmarker, 15).AddComment _
    "New QTO (import skipped) = " & Format(dat(qtorowmarker, 6), "###,##0") & " " & Cells(datarowmarker, 14).Value
End If

dat(qtorowmarker, 1) = "skipped & flagged"
Call finishclick

End Sub

Private Sub dataloop()

For x = datarowmarker To lastrow Step 1
    For i = LBound(dat, 1) To UBound(dat, 1)
        If dat(i, 4) = Cells(x, 12).Value And dat(i, 1) = "" Then
            Cells(x, 12).Select
            matchcount = matchcount + 1
            lmatchcount.Caption = matchcount
            litemdescription.Caption = dat(i, 4)
            lcontractitem.Caption = dat(i, 3)
            luniformatitem.Caption = dat(i, 2)
            lcurrenttotal.Caption = Format(Cells(x, 16).Value, "$#,##0")
            lcurrentquantity.Caption = Format(Cells(x, 15).Value, "###,##0") & " " & Cells(x, 14).Value
            lnewtotal.Caption = Format(dat(i, 6) * Cells(x, 13).Value, "$#,##0")
            lnewquantity.Caption = Format(dat(i, 6), "###,##0") & " " & Cells(x, 14).Value
            If IsNumeric(Cells(x, 16).Value) Then
                ldeltatotal.Caption = Format(dat(i, 6) * Cells(x, 13).Value - Cells(x, 16).Value, "$#,##0")
            Else: ldeltatotal.Caption = Format(dat(i, 6) * Cells(x, 13).Value, "$#,##0")
            End If
            If IsNumeric(Cells(x, 15).Value) Then
                ldeltaquantity.Caption = Format(dat(i, 6) - Cells(x, 15).Value, "###,##0") & " " & Cells(x, 14).Value
            Else: ldeltaquantity.Caption = Format(dat(i, 6), "###,##0") & " " & Cells(x, 14).Value
            End If
            datarowmarker = x
            qtorowmarker = i
            Exit Sub
        End If
    Next
Next x

dataloopclosed = True
lineitemframe.Caption = "New Line Item"
Call qtoloop

End Sub

Sub qtoloop()

For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 1) = "" Then
        newcount = newcount + 1
        lnewcount.Caption = newcount
        litemdescription.Caption = dat(i, 4)
        lcontractitem.Caption = dat(i, 3)
        luniformatitem.Caption = dat(i, 2)
        lcurrenttotal.Caption = "N/A"
        lcurrentquantity.Caption = "N/A"
        lnewtotal.Caption = "N/A"
        lnewquantity.Caption = Format(dat(i, 6), "###,##0") & " " & dat(i, 5)
        ldeltatotal.Caption = "N/A"
        ldeltaquantity.Caption = "N/A"
        qtorowmarker = i
        Exit Sub
    End If
Next
Me.Hide
Sheets("Dashboard").Activate
Range("A1").Select
End Sub

Sub finishclick()


rng.Value = dat
Me.AddProgress 1

If dataloopclosed = False Then
    datarowmarker = datarowmarker + 1
    Call dataloop
Else
    qtorowmarker = qtorowmarker + 1
    Call qtoloop
End If

End Sub

Sub qtolineimport()

End Sub

