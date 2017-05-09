Attribute VB_Name = "Auserforms"
Public pb As ProgressBar

Sub progressIndicator_Begin(jobName As String)
    Set pb = New ProgressBar
    pb.Initialize jobName, 100
End Sub

Sub progressIndicator_End()
    pb.Hide
    Set pb = Nothing
End Sub

Sub workbookSelect()
workbookselectuserform.Show
End Sub

Sub copylineitems(destinationwb As String)

Dim wbThis As Workbook
Set wbThis = ActiveWorkbook
Dim rng As range
Dim copyrng As String

Set rng = Selection
nLastRow = rng.Rows.count + rng.Row - 1
nFirstRow = rng.Row

copyrng = "$A$" & nFirstRow & ":$N$" & nLastRow

Application.Workbooks(destinationwb).Worksheets("Data").Activate

Dim rw As Integer
If range("L6").Value = "" And range("L7").Value = "" Then
    rw = 6
Else: rw = range("L6").End(xlDown).Row + 1
End If

Rows(rw + 1 & ":" & rw + rng.Rows.count).Insert Shift:=xlDown

wbThis.Worksheets("Data").Activate
range(copyrng).Copy
Application.DisplayAlerts = False
Application.Workbooks(destinationwb).Worksheets("Data").Activate
range("A" & rw + 1).PasteSpecial _
        Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

Application.CutCopyMode = False

range("Q" & rw + 1 & ":" & "AB" & rw + rng.Rows.count).ClearContents

If rw = 6 Then
    Rows(6).Delete
End If

wbThis.Worksheets("Data").Activate
Application.DisplayAlerts = True
MsgBox (rng.Rows.count & " Line Items Copied!")
End Sub
