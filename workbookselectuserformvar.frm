VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} workbookselectuserformvar 
   Caption         =   "Select Variance Estimate"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   OleObjectBlob   =   "workbookselectuserformvar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "workbookselectuserformvar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If workbookselectuserformvar.ListBox1.ListIndex < 0 Then
        MsgBox ("either select a workbook or close the dialog box")
    Else
        Call detailVariance(workbookselectuserformvar.ListBox1.Value)
    End If
    
End Sub

Private Sub UserForm_Initialize()

workbookselectuserformvar.ListBox1.MultiSelect = False

If ActiveSheet.Name <> "dashboard" Then
    MsgBox ("error")
    Exit Sub
Else
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If sheetExists("Data", wb) = True And wb.Name <> Application.ActiveWorkbook.Name Then
            workbookselectuserformvar.ListBox1.AddItem wb.Name
        End If
    Next wb
End If

End Sub


Function sheetExists(sheetToFind As String, wb As Workbook) As Boolean
    sheetExists = False
    For Each Sheet In wb.Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function


