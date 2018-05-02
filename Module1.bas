Attribute VB_Name = "Module1"
Sub setzones()

Application.Worksheets("Data").Activate

'HIDE UNUSED COLUMNS
If Range("name_Z2").Value = "" Then
    Sheets("data").Columns("R").EntireColumn.Hidden = True
    Sheets("data").Columns("AD").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("R").EntireColumn.Hidden = False
    Sheets("data").Columns("AD").EntireColumn.Hidden = False
End If

If Range("name_Z3").Value = "" Then
    Sheets("data").Columns("S").EntireColumn.Hidden = True
    Sheets("data").Columns("AE").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("S").EntireColumn.Hidden = False
    Sheets("data").Columns("AE").EntireColumn.Hidden = False
End If

If Range("name_Z4").Value = "" Then
    Sheets("data").Columns("T").EntireColumn.Hidden = True
    Sheets("data").Columns("AF").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("T").EntireColumn.Hidden = False
    Sheets("data").Columns("AF").EntireColumn.Hidden = False
End If

If Range("name_Z5").Value = "" Then
    Sheets("data").Columns("U").EntireColumn.Hidden = True
    Sheets("data").Columns("AG").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("U").EntireColumn.Hidden = False
    Sheets("data").Columns("AG").EntireColumn.Hidden = False
End If

If Range("name_Z6").Value = "" Then
    Sheets("data").Columns("V").EntireColumn.Hidden = True
    Sheets("data").Columns("AH").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("V").EntireColumn.Hidden = False
    Sheets("data").Columns("AH").EntireColumn.Hidden = False
End If

If Range("name_Z7").Value = "" Then
    Sheets("data").Columns("W").EntireColumn.Hidden = True
    Sheets("data").Columns("AI").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("W").EntireColumn.Hidden = False
    Sheets("data").Columns("AI").EntireColumn.Hidden = False
End If

If Range("name_Z8").Value = "" Then
    Sheets("data").Columns("X").EntireColumn.Hidden = True
    Sheets("data").Columns("AJ").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("X").EntireColumn.Hidden = False
    Sheets("data").Columns("AJ").EntireColumn.Hidden = False
End If

If Range("name_Z9").Value = "" Then
    Sheets("data").Columns("Y").EntireColumn.Hidden = True
    Sheets("data").Columns("AK").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("Y").EntireColumn.Hidden = False
    Sheets("data").Columns("AK").EntireColumn.Hidden = False
End If

If Range("name_Z10").Value = "" Then
    Sheets("data").Columns("Z").EntireColumn.Hidden = True
    Sheets("data").Columns("AL").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("Z").EntireColumn.Hidden = False
    Sheets("data").Columns("AL").EntireColumn.Hidden = False
End If

If Range("name_Z11").Value = "" Then
    Sheets("data").Columns("AA").EntireColumn.Hidden = True
    Sheets("data").Columns("AM").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("AA").EntireColumn.Hidden = False
    Sheets("data").Columns("AM").EntireColumn.Hidden = False
End If

If Range("name_Z12").Value = "" Then
    Sheets("data").Columns("AB").EntireColumn.Hidden = True
    Sheets("data").Columns("AN").EntireColumn.Hidden = True
Else
    Sheets("data").Columns("AB").EntireColumn.Hidden = False
    Sheets("data").Columns("AN").EntireColumn.Hidden = False
End If

End Sub
