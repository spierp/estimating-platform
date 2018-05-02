Attribute VB_Name = "externalReferences"
Sub FindLinksInValidationWorkbook()

Dim rCell As Range
Dim sDvForm As String
Dim counter As Integer
Dim ws As Worksheet

Application.ScreenUpdating = False

'creates a worksheet called external links
Set ws = ThisWorkbook.Sheets.Add(After:= _
         ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
ws.Name = "external links"
Sheets("external links").Cells(1, 1).Value = "cell address"
Sheets("external links").Cells(1, 2).Value = "worksheet"


'unhide all worksheets
For Each ws In Worksheets
    ws.Visible = True
Next

'set a counter
counter = 2

'loop through all the worksheets and find the offending references
For Each ws In Worksheets
    ws.Select
    Cells.Select

    For Each rCell In ActiveSheet.UsedRange.Cells
        'Store the Formula1 property if there is one
        On Error Resume Next
            sDvForm = ""
            sDvForm = rCell.Validation.Formula1
        On Error GoTo 0

        'If Formula1 has a bracket, it’s a good candidate
        'for containing an external link
        If InStr(1, sDvForm, "]") > 0 Then
            Sheets("external links").Cells(counter, 1).Value = rCell.Address
            Sheets("external links").Cells(counter, 2).Value = ActiveSheet.Name

            'this gets the formula and the external file, but i don't need it. uncomment it if you find it useful
            'Sheets("external links").Cells(counter, 3).value = rCell.Validation.Formula1

            counter = counter + 1

        End If

        'selects one cell in the worksheet after searching through all the cells
        Range("A1").Select
    Next rCell

Next

Application.ScreenUpdating = True
End Sub

