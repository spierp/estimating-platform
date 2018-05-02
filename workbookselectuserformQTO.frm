VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} workbookselectuserformQTO 
   Caption         =   "Select QTO Table"
   ClientHeight    =   8640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   OleObjectBlob   =   "workbookselectuserformQTO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "workbookselectuserformQTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim qamode As Boolean

Private Sub CommandButton1_Click()
    If workbookselectuserformQTO.ListBox1.ListIndex < 0 Then
        MsgBox ("Please select the QTO source workbook")
    ElseIf OptionButtonLevel = False And OptionButtonZone = False Then
        MsgBox ("Please select column format. Either group QTOs by 'zone' or 'level'.")
    Else
        workbookselectuserformQTO.Hide
        If qamode = True Then
            Set qtoForm = New qtouserform
                qtoForm.Initialize (workbookselectuserformQTO.ListBox1.Value)
            Else
            Set qtoDog = New ProgressBarHotdog
                qtoDog.Initialize (workbookselectuserformQTO.ListBox1.Value)
        End If
    End If
End Sub

Private Sub OptionButtonHD_Click()
qamode = False
End Sub

Private Sub OptionButtonLevel_Click()
zonemode = False
End Sub

Private Sub OptionButtonQA_Click()
qamode = True
End Sub

Private Sub OptionButtonZone_Click()
zonemode = True
End Sub

Private Sub UserForm_Initialize()

workbookselectuserformQTO.ListBox1.MultiSelect = False
qamode = True
If ActiveSheet.Name <> "dashboard" Then
    MsgBox ("error")
    Exit Sub
Else
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If sheetExists("dynamo-export", wb) = True And wb.Name <> Application.ActiveWorkbook.Name Then
            workbookselectuserformQTO.ListBox1.AddItem wb.Name
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


