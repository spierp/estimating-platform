VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} wizard 
   Caption         =   "Estimate Setup Wizard"
   ClientHeight    =   14775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20235
   OleObjectBlob   =   "wizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "wizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long, constructiontypes As String, dat As Variant, x As Integer, selectedlistitems As String, _
datmarker As Variant, rng As Range, rngmarker As Range, activelist As String, activelistcolumn As Integer, _
destwb As Workbook, sFName As String, dbwb As Workbook, rowcount As Integer, linecopyrng As String, _
qtycopyrng As String, zonecount As Integer, zone As Integer, lastrow As Integer, rw As Integer

Private Sub UserForm_Initialize()

    Me.constructionscopesprimlistbox.List = Range("constructionscopeprim").Value
    Me.archsystemslistbox.List = Range("archsys").Value
    Me.MEPFsystemslistbox.List = Range("MEPFsys").Value
    Me.logisticslistbox.List = Range("logistics").Value
    Me.designlistbox.List = Range("design_disciplines").Value
    Me.gicsindustrygrouplistbox.List = Range("GICS.IndustryGroup").Value
    Me.spaceslistbox.List = Range("spaces").Value
    Me.formslistbox.List = Range("entitiesbyform").Value
    Me.functionslistbox1.List = Range("entitiesbyfunction").Value
    Me.MultiPage1.Value = 0
    
    Me.zonecountbox.AddItem "1"
    Me.zonecountbox.AddItem "2"
    Me.zonecountbox.AddItem "3"
    Me.zonecountbox.AddItem "4"
    Me.zonecountbox.AddItem "5"
    Me.zonecountbox.AddItem "6"
    Me.zonecountbox.AddItem "7"
    Me.zonecountbox.AddItem "8"
    Me.zonecountbox.AddItem "9"
    Me.zonecountbox.AddItem "10"
    Me.zonecountbox.AddItem "11"
    Me.zonecountbox.AddItem "12"
    
End Sub


Private Sub dbb_Click()
For i = 0 To designlistbox.ListCount - 1
    If designlistbox.List(i) = "33-21_31_17_21_Fire_Protection_Engineering" Or designlistbox.List(i) = "33-21_31_99_21_21_Alarm_and_Detection_Engineering" Then
        designlistbox.Selected(i) = True
    End If
Next

End Sub
Private Sub db_Click()

For i = 0 To designlistbox.ListCount - 1
    If designlistbox.List(i) = "33-21_11_Architecture" Or designlistbox.List(i) = "33-21_31_14_Structural_Engineering" _
    Or designlistbox.List(i) = "33-21_31_17_31_HVAC_Engineering" Or designlistbox.List(i) = "33-21_31_17_11_Plumbing_Engineering" _
    Or designlistbox.List(i) = "33-21_31_17_21_Fire_Protection_Engineering" Or designlistbox.List(i) = "33-21_31_21_Electrical_Engineering" _
    Or designlistbox.List(i) = "33-21_31_99_21_21_Alarm_and_Detection_Engineering" Then
        designlistbox.Selected(i) = True
    End If
Next

End Sub
Private Sub dbmep_Click()

For i = 0 To designlistbox.ListCount - 1
    If designlistbox.List(i) = "33-21_31_17_31_HVAC_Engineering" Or designlistbox.List(i) = "33-21_31_17_11_Plumbing_Engineering" _
    Or designlistbox.List(i) = "33-21_31_17_21_Fire_Protection_Engineering" Or designlistbox.List(i) = "33-21_31_21_Electrical_Engineering" _
    Or designlistbox.List(i) = "33-21_31_99_21_21_Alarm_and_Detection_Engineering" Then
        designlistbox.Selected(i) = True
    End If
Next

End Sub

Private Sub designclearcommandbutton_Click()

For i = 0 To designlistbox.ListCount - 1
    designlistbox.Selected(i) = False
Next

End Sub

Private Sub functionslistbox1_Click()

For i = 0 To functionslistbox1.ListCount - 1
    If functionslistbox1.Selected(i) = True Then
        Me.functionslistbox2.List = Range("Omni" & functionslistbox1.List(i)).Value
    End If
Next

End Sub

Private Sub gicsindustrygrouplistbox_Click()

For i = 0 To gicsindustrygrouplistbox.ListCount - 1
    If gicsindustrygrouplistbox.Selected(i) = True Then
    Me.gicsindustrylistbox.List = Range("GICS." & gicsindustrygrouplistbox.List(i)).Value
    End If
Next

End Sub


Private Sub P1next_Click()
    Me.MultiPage1.Value = 1
    Me.pagelabel1.ForeColor = RGB(160, 160, 160) 'gray
    Me.pagelabel2.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel1.Font.Underline = False
    Me.pagelabel2.Font.Underline = True
    
    constructiontypes = GetSelectedItems(Me.constructionscopesprimlistbox)
    If constructiontypes Like "*01_Ground-up_Cold_Shell_Construction*" Then
        dat = Range("SGroundup_Cold_Shell_Construction").Value
        For x = LBound(dat, 1) To UBound(dat, 1)
            For i = 0 To spaceslistbox.ListCount - 1
                If dat(x, 1) = spaceslistbox.List(i) Then
                    spaceslistbox.Selected(i) = True
                End If
            Next
        Next
    End If
    If constructiontypes Like "*02_Ground-up_Core_&_Shell_Warm-Up*" Then
        dat = Range("SGroundup_Core_and_Shell_WarmUp").Value
        For x = LBound(dat, 1) To UBound(dat, 1)
            For i = 0 To spaceslistbox.ListCount - 1
                If dat(x, 1) = spaceslistbox.List(i) Then
                    spaceslistbox.Selected(i) = True
                End If
            Next
        Next
    End If
    If constructiontypes Like "*03_Core_&_Shell_Renovation*" Then
        dat = Range("SCore_and_Shell_Renovation").Value
        For x = LBound(dat, 1) To UBound(dat, 1)
            For i = 0 To spaceslistbox.ListCount - 1
                If dat(x, 1) = spaceslistbox.List(i) Then
                    spaceslistbox.Selected(i) = True
                End If
            Next
        Next
    End If
    If constructiontypes Like "*04_Interior_Construction*" Then
        dat = Range("SInterior_Construction").Value
        For x = LBound(dat, 1) To UBound(dat, 1)
            For i = 0 To spaceslistbox.ListCount - 1
                If dat(x, 1) = spaceslistbox.List(i) Then
                    spaceslistbox.Selected(i) = True
                End If
            Next
        Next
    End If
    If constructiontypes Like "*05_Interior_MEPF*" Then
        dat = Range("SInterior_MEPF").Value
        For x = LBound(dat, 1) To UBound(dat, 1)
            For i = 0 To spaceslistbox.ListCount - 1
                If dat(x, 1) = spaceslistbox.List(i) Then
                    spaceslistbox.Selected(i) = True
                End If
            Next
        Next
    End If
    If constructiontypes Like "*06_Site_Work*" Then
        dat = Range("SSite_Work").Value
        For x = LBound(dat, 1) To UBound(dat, 1)
            For i = 0 To spaceslistbox.ListCount - 1
                If dat(x, 1) = spaceslistbox.List(i) Then
                    spaceslistbox.Selected(i) = True
                End If
            Next
        Next
    End If

End Sub

Private Sub P2prev_Click()
    Me.MultiPage1.Value = 0
    Me.pagelabel1.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel2.ForeColor = RGB(152, 152, 152) 'gray
    Me.pagelabel2.Font.Underline = False
    Me.pagelabel1.Font.Underline = True
End Sub
Private Sub P2next_Click()
    Me.MultiPage1.Value = 2
    Me.pagelabel2.ForeColor = RGB(152, 152, 152) 'gray
    Me.pagelabel3.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel2.Font.Underline = False
    Me.pagelabel3.Font.Underline = True
End Sub

Private Sub P3prev_Click()
    Me.MultiPage1.Value = 1
    Me.pagelabel3.ForeColor = RGB(152, 152, 152) 'gray
    Me.pagelabel2.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel3.Font.Underline = False
    Me.pagelabel2.Font.Underline = True
End Sub

Private Sub P3next_Click()
    Me.MultiPage1.Value = 3
    Me.pagelabel3.ForeColor = RGB(152, 152, 152) 'gray
    Me.pagelabel4.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel3.Font.Underline = False
    Me.pagelabel4.Font.Underline = True
End Sub

Private Sub P4prev_Click()
    Me.MultiPage1.Value = 2
    Me.pagelabel4.ForeColor = RGB(152, 152, 152) 'gray
    Me.pagelabel3.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel4.Font.Underline = False
    Me.pagelabel3.Font.Underline = True
End Sub

Private Sub P4next_Click()
    Me.MultiPage1.Value = 4
    Me.pagelabel4.ForeColor = RGB(152, 152, 152) 'gray
    Me.pagelabel5.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel4.Font.Underline = False
    Me.pagelabel5.Font.Underline = True
End Sub

Private Sub P5prev_Click()
    Me.MultiPage1.Value = 3
    Me.pagelabel5.ForeColor = RGB(152, 152, 152) 'gray
    Me.pagelabel4.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel5.Font.Underline = False
    Me.pagelabel4.Font.Underline = True
End Sub

Private Sub P5next_Click()
    Me.MultiPage1.Value = 5
    Me.pagelabel5.ForeColor = RGB(152, 152, 152) 'gray
    Me.pagelabel4.ForeColor = RGB(255, 0, 0) 'red
    Me.pagelabel5.Font.Underline = False
    Me.pagelabel4.Font.Underline = True
End Sub

'LINKS
Private Sub GICSlabel_Click()
    ThisWorkbook.FollowHyperlink Address:="https://en.wikipedia.org/wiki/Global_Industry_Classification_Standard"
End Sub
Private Sub omni11url_Click()
    ThisWorkbook.FollowHyperlink Address:="http://www.omniclass.org/"
End Sub
Private Sub omni12url_Click()
    ThisWorkbook.FollowHyperlink Address:="http://www.omniclass.org/"
End Sub

Private Sub omni13url_Click()
    ThisWorkbook.FollowHyperlink Address:="http://www.omniclass.org/"
End Sub

Private Sub omni33url_Click()
    ThisWorkbook.FollowHyperlink Address:="http://www.omniclass.org/"
End Sub

Private Sub wikisearch_Click()
    ThisWorkbook.FollowHyperlink Address:="https://www.wikipedia.org/"
End Sub

Private Sub taglineitemsbutton_Click()

If Me.zonecountbox.Text = "" Then
    MsgBox ("please select number of zones")
Else

Worksheets("clipboard").Cells.Clear

lastrow = Worksheets("lineitems").ListObjects("lineitemsTable").ListColumns(1).Range.Rows.Count + 2

zonecount = Me.zonecountbox.Text

Set rng = Worksheets("lineitems").Range("A3:AC" & lastrow)
Set rngmarker = Worksheets("lineitems").Range("AC3:AC" & lastrow)

dat = rng.Value
datmarker = rngmarker.Value

'TAG LINE ITEMS
Call taglineitems(spaceslistbox, 7)
Call taglineitems(constructionscopesprimlistbox, 12)
Call taglineitems(constructionscopesseclistbox, 13)
Call taglineitems(archsystemslistbox, 14)
Call taglineitems(MEPFsystemslistbox, 15)
Call taglineitems(logisticslistbox, 19)
Call taglineitems(functionslistbox2, 24)
Call taglineitems(formslistbox, 25)
Call taglineitems(designlistbox, 26)


'TAG LINE ITEMS WITH RELATIONSHIPS - LOOP
For x = LBound(datmarker, 1) To UBound(datmarker, 1)
    If datmarker(x, 1) = "X" Then
        For i = LBound(dat, 1) To UBound(dat, 1)
            If dat(i, 16) Like "*" & dat(x, 1) & "*" Then
                datmarker(i, 1) = "X"
            End If
        Next i
    End If
Next x

rngmarker.Value = datmarker

Worksheets("lineitems").ListObjects("lineitemsTable").Range.AutoFilter Field:=29, Criteria1:="X", Operator:=xlFilterValues
    
    Worksheets("lineitems").ListObjects("lineitemsTable").Range.Copy
    Worksheets("clipboard").Range("A1").PasteSpecial _
    Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
    xlNone, SkipBlanks:=True, Transpose:=False
    
    Application.CutCopyMode = False
    
Worksheets("lineitems").ListObjects("lineitemsTable").AutoFilter.ShowAllData

Worksheets("lineitems").ListObjects("lineitemsTable").ListColumns(29).Range.ClearContents

Me.Hide
Worksheets("clipboard").Activate

Set dbwb = ActiveWorkbook
Set destwb = Workbooks.Open("\\HD-DATA.HDCCo.com\UserGroup$\PreCon\Estimating Master\10 Templates\Estimates\Rainier Estimate Template.xlsm")

Range("prim_div_qty").Value = "1"
'Range("dur_qty").Value = "1"
'Range("count_qty").Value = "1"

zone = 1
For x = 1 To zonecount Step 1
    Range("name_z" & zone).Value = "name" & zone
    Range("description_z" & zone).Value = "desc_" & zone
    Range("prim_div_qty_z" & zone).Value = "1"
    zone = zone + 1
Next x

sFName = Application.GetSaveAsFilename _
(InitialFileName:="Estimate File Name", _
    FileFilter:="Excel files (*.xlsm), *.xlsm", _
    Title:="Set location and filename of your new estimate")

If sFName <> "False" Then
    destwb.SaveAs sFName
Else
    MsgBox ("New Estimate Cancelled")
    Exit Sub
End If

Set destwb = ActiveWorkbook

dbwb.Sheets("clipboard").Activate
rowcount = Cells(Rows.Count, "A").End(xlUp).Row

linecopyrng = "$A$2:$J$" & rowcount
qtycopyrng = "$K$2:$K$" & rowcount

destwb.Worksheets("Data").Activate

If Range("L6").Value = "" And Range("L7").Value = "" Then
    rw = 6
Else: rw = Range("L6").End(xlDown).Row + 1
End If

Rows(rw + 1 & ":" & rw + rowcount).Insert Shift:=xlDown

dbwb.Worksheets("clipboard").Activate

Range(linecopyrng).Copy
Application.DisplayAlerts = False
destwb.Worksheets("Data").Activate
Range("E" & rw + 1).PasteSpecial _
        Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

Application.CutCopyMode = False

'Range("Q" & rw + 1 & ":" & "AB" & rw + rng.Rows.Count).ClearContents

dbwb.Worksheets("clipboard").Activate
Range(qtycopyrng).Copy
destwb.Worksheets("Data").Activate
Range("Q" & rw + 1).PasteSpecial _
        Paste:=xlPasteFormulasAndNumberFormats, Operation:= _
        xlNone, SkipBlanks:=False, Transpose:=False

Application.CutCopyMode = False

If rw = 6 Then
    Rows(6).Delete
End If

Application.AutoCorrect.AutoFillFormulasInLists = False

'SET CORRECT NUMBER OF COLUMNS
Application.Run "'" & destwb.Name & "'!setzones"

'CONVERT FORMULA VALUE TO PRACTICAL CELL FORMULA
Dim cellformula As String

For x = rw + rowcount + 1 To 7 Step -1
    If Cells(x, 17).Value = "proratebasedonprimediv" Then
        
        Range("M" & x).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Range("Q" & x & ":" & "AB" & x).NumberFormat = "0%"
        
        zone = 1
        For i = 17 To 16 + zonecount Step 1
            Cells(x, i).formula = "=prim_div_qty_z" & zone & "/prim_div_qty"
            zone = zone + 1
        Next i
        
    ElseIf Cells(x, 17).Value = "primdivqty" Then
        
        zone = 1
        For i = 17 To 16 + zonecount Step 1
            Cells(x, i).formula = "=prim_div_qty_z" & zone
            zone = zone + 1
        Next i
        
    ElseIf Cells(x, 17).Value = "durqty" Then
        
        zone = 1
        For i = 17 To 16 + zonecount Step 1
            Cells(x, i).formula = "=dur_z" & zone
            zone = zone + 1
        Next i
        
    ElseIf Cells(x, 17).Value Like "*" & "primdivqtyformula" & "*" Then
        zone = 1
        cellformula = Cells(x, 17).Value
        For i = 17 To 16 + zonecount Step 1
            Cells(x, i).formula = "=" & Replace(cellformula, "primdivqtyformula", "prim_div_qty_z" & zone)
            zone = zone + 1
        Next i
    ElseIf Cells(x, 17).Value Like "VLOOKUP" & "*" Then
        zone = 1
        cellformula = Cells(x, 17).Value
        For i = 17 To 16 + zonecount Step 1
            Cells(x, i).formula = "=" & Replace(cellformula, "zzz", zone + 12)
            zone = zone + 1
        Next i
    Else
    End If
Next x

Application.AutoCorrect.AutoFillFormulasInLists = True

destwb.Worksheets("dashboard").Activate

Application.DisplayAlerts = True
MsgBox (rowcount - 1 & " Database Line Items Entered!")

destwb.Save
dbwb.Close False
End If

End Sub
Sub taglineitems(activelist As MSForms.ListBox, activelistcolumn As Integer)

For i = 0 To activelist.ListCount - 1
    If activelist.Selected(i) = True Then
        For x = LBound(dat, 1) To UBound(dat, 1)
            If dat(x, activelistcolumn) Like "*" & activelist.List(i) & "*" Or dat(x, activelistcolumn) = "Z_All" Then
                datmarker(x, 1) = "X"
            End If
        Next
    ElseIf activelist.Selected(i) = False Then
        For x = LBound(dat, 1) To UBound(dat, 1)
            If dat(x, activelistcolumn) Like "*" & "<>" & activelist.List(i) & "*" Then
                datmarker(x, 1) = "X"
            End If
        Next
    End If
Next

End Sub
Public Function GetSelectedItems(lBox As MSForms.ListBox) As String
'returns an array of selected items in a ListBox
Dim tmpArray() As Variant
Dim i As Integer
Dim selCount As Integer
    selCount = -1
    '## Iterate over each item in the ListBox control:
    For i = 0 To lBox.ListCount - 1
        '## Check to see if this item is selected:
        If lBox.Selected(i) = True Then
            '## If this item is selected, then add it to the array
            selCount = selCount + 1
            ReDim Preserve tmpArray(selCount)
            tmpArray(selCount) = lBox.List(i)
        End If
    Next

    If selCount = -1 Then
        '## If no items were selected, return an empty string
        GetSelectedItems = "" ' or "No items selected", etc.
    Else:
        '## Otherwise, return the array of items as a string,
        '   delimited by commas
        GetSelectedItems = Join(tmpArray, ", ")
    End If
End Function




