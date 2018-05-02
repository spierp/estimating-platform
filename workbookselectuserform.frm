VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} workbookselectuserform 
   Caption         =   "Select Destination Workbook"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6120
   OleObjectBlob   =   "workbookselectuserform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "workbookselectuserform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If workbookselectuserform.ListBox1.ListIndex < 0 Then
        MsgBox ("either select a workbook or close the dialog box")
    Else
        Call copylineitems(workbookselectuserform.ListBox1.Value)
        workbookselectuserform.Hide
        Set workbookselectuserform = Nothing
    End If
    
End Sub

Private Sub UserForm_Initialize()

workbookselectuserform.ListBox1.MultiSelect = False

If ActiveSheet.Name <> "Data" Or ActiveCell.Row < 6 Or _
ActiveCell.Row > Worksheets("Data").ListObjects("dataTable").ListColumns(5).Range.Rows.Count Then
    MsgBox ("Please select relevant line item(s) first")
    Exit Sub
'    workbookselectuserform.Hide
'    Set workbookselectuserform = Nothing
Else
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If sheetExists("Data", wb) = True And wb.Name <> Application.ActiveWorkbook.Name Then
            workbookselectuserform.ListBox1.AddItem wb.Name
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
