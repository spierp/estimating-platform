VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} progressBarHotdog 
   Caption         =   "QTO Import - Hotdog Mode"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9015
   OleObjectBlob   =   "progressBarHotdog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBarHotdog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim zonecount As Integer, wbThis As String, i As Integer, max As Long, dat As Variant, varTemp As Variant, rng As Range, lastrow As Integer
Dim matchcount As Integer, completed As Integer, newcount As Integer, rw As Integer

Dim progress As Double, maxProgress As Double, maxWidth As Long, startTime As Double, rc As Integer

Public Sub Initialize(qtowb As String)
wbThis = ActiveWorkbook.Name

Call dynamo(qtowb)
If fatalerror = True Then
    Application.Workbooks(wbThis).Worksheets("Dashboard").Activate
    Range("A1").Select
    Me.Hide
    Exit Sub
End If

Application.Workbooks(qtowb).Worksheets("QTO").Activate

zonecount = WorksheetFunction.CountA(Range("G1:R1"))
max = Cells(Rows.Count, 4).End(xlUp).Row - 1
MsgBox ("max is" & max)
maxProgress = max:  maxWidth = lBar.Width:    lBar.Width = 0
lProgress.Caption = "0"
lcount.Caption = max
lzones.Caption = zonecount

Range("A1").CurrentRegion.Select
Set rng = Selection.Cells
dat = rng.Value

Application.Workbooks(wbThis).Worksheets("Data").Activate
Cells.clearcomments

Me.Show False

lastrow = Cells(Rows.Count, 1).End(xlUp).Row - 1

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

Private Sub dataloop()

For x = 6 To lastrow Step 1
    For i = LBound(dat, 1) To UBound(dat, 1)
        If dat(i, 4) = Cells(x, 12).Value And dat(i, 1) = "" Then
            Cells(x, 12).Select
            
            'update form
            matchcount = matchcount + 1
            lmatchcount.Caption = matchcount
            litemdescription.Caption = dat(i, 4)
            lcontractitem.Caption = dat(i, 3)
            luniformatitem.Caption = dat(i, 2)
            Me.AddProgress 1
            
            'insert qto
            Cells(x, 14).Value = dat(i, 5)
            Cells(x, 15).AddComment _
                "Previous QTO = " & Format(Cells(x, 15).Value, "###,##0") & " " & Cells(x, 14).Value
            
            If zonecount = 1 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7)))
            ElseIf zonecount = 2 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8)))
            ElseIf zonecount = 3 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9)))
            ElseIf zonecount = 4 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10)))
            ElseIf zonecount = 5 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11)))
            ElseIf zonecount = 6 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12)))
            ElseIf zonecount = 7 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13)))
            ElseIf zonecount = 8 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14)))
            ElseIf zonecount = 9 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15)))
            ElseIf zonecount = 10 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16)))
            ElseIf zonecount = 11 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17)))
            ElseIf zonecount = 12 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18)))
            End If
            
            Range("Q" & x).Resize(LBound(varTemp), UBound(varTemp)) = Application.Transpose(varTemp)
            dat(i, 1) = "imported"

        End If
    Next
Next x

rng.Value = dat

rc = rc - matchcount

Call qtoloop

End Sub

Sub qtoloop()

rw = Range("L6").End(xlDown).Row + 1
Range("A" & rw & ":A" & rw + max - matchcount).EntireRow.Insert

For i = LBound(dat, 1) To UBound(dat, 1)
    If dat(i, 1) = "" Then
        
        'update form
        newcount = newcount + 1
        lnewcount.Caption = newcount
        litemdescription.Caption = dat(i, 4)
        lcontractitem.Caption = dat(i, 3)
        luniformatitem.Caption = dat(i, 2)
        Me.AddProgress 1
        'insert qto
        rw = rw + 1
            If zonecount = 1 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7)))
            ElseIf zonecount = 2 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8)))
            ElseIf zonecount = 3 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9)))
            ElseIf zonecount = 4 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10)))
            ElseIf zonecount = 5 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11)))
            ElseIf zonecount = 6 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12)))
            ElseIf zonecount = 7 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13)))
            ElseIf zonecount = 8 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14)))
            ElseIf zonecount = 9 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15)))
            ElseIf zonecount = 10 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16)))
            ElseIf zonecount = 11 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17)))
            ElseIf zonecount = 12 Then
                varTemp = Application.Index(dat, i, Application.Transpose(Array(7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18)))
            End If
            
        Range("Q" & rw).Resize(LBound(varTemp), UBound(varTemp)) = Application.Transpose(varTemp)
        
        Cells(rw, 14).Value = dat(i, 5)
        Cells(rw, 12).Value = dat(i, 4)
        Cells(rw, 10).Value = dat(i, 3)
        Cells(rw, 9).Value = dat(i, 2)

        dat(i, 1) = "imported"
        
    End If
Next
rng.Value = dat

Me.Hide
Sheets("Dashboard").Activate
Range("A1").Select
End Sub
