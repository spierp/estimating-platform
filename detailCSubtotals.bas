Attribute VB_Name = "detailCSubtotals"
Sub createSubTotals()
Attribute createSubTotals.VB_ProcData.VB_Invoke_Func = " \n14"
pb.Repaint
'pb 20% complete

'CREATE SUBTOTALS
pb.AddCaption "Calculating Subtotals..."

    Dim zonenumber As Integer
    zonenumber = WorksheetFunction.CountA(range("Q6:AY6")) / 2
    range("A6").CurrentRegion.Select
    
    If zonenumber = 12 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 29, 30 _
            , 31, 32, 33, 34, 35, 36, 37, 38, 39, 40), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 29, 30 _
                , 31, 32, 33, 34, 35, 36, 37, 38, 39, 40), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 29, 30 _
                , 31, 32, 33, 34, 35, 36, 37, 38, 39, 40), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
    
    ElseIf zonenumber = 11 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 28, 29 _
            , 30, 31, 32, 33, 34, 35, 36, 37, 38), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, FFunction:=xlSum, TotalList:=Array(16, 28, 29 _
                , 30, 31, 32, 33, 34, 35, 36, 37, 38), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 28, 29 _
                , 30, 31, 32, 33, 34, 35, 36, 37, 38), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 10 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 27, 28 _
            , 29, 30, 31, 32, 33, 34, 35, 36), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 27, 28 _
                , 29, 30, 31, 32, 33, 34, 35, 36), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 27, _
                28, 29, 30, 31, 32, 33, 34, 35, 36), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 9 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 26, 27 _
            , 28, 29, 30, 31, 32, 33, 34), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 26, 27 _
                , 28, 29, 30, 31, 32, 33, 34), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 26, 27 _
                , 28, 29, 30, 31, 32, 33, 34), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 8 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 25, 26 _
            , 27, 28, 29, 30, 31, 32), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 25, 26 _
                , 27, 28, 29, 30, 31, 32), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 25, 26 _
                , 27, 28, 29, 30, 31, 32), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 7 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 24, 25 _
            , 26, 27, 28, 29, 30), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 24, 25 _
                , 26, 27, 28, 29, 30), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 24, 25 _
                , 26, 27, 28, 29, 30), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 6 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 23, 24 _
            , 25, 26, 27, 28), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 23, 24 _
                , 25, 26, 27, 28), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 23, 24 _
                , 25, 26, 27, 28), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 5 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 22, 23 _
            , 24, 25, 26), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
        range("A6").CurrentRegion.Select
        Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 22, 23 _
            , 24, 25, 26), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
        range("A6").CurrentRegion.Select
        Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 22, 23 _
            , 24, 25, 26), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 4 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 21, 22 _
            , 23, 24), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 21, 22 _
                , 23, 24), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
        range("A6").CurrentRegion.Select
        Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 21, 22 _
            , 23, 24), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 3 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 20, 21 _
            , 22), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 20, 21 _
                , 22), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 20, 21 _
                , 22), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 2 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16, 19, 20), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16, 19, 20), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16, 19, 20), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        
    ElseIf zonenumber = 1 Then
        Selection.Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(16), Replace:=False, PageBreaks:=False, _
            SummaryBelowData:=True
        If Worksheets("Dashboard").range("subtotals_L2") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=2, Function:=xlSum, TotalList:=Array(16), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
        End If
        If Worksheets("Dashboard").range("subtotals_L3") = "Yes" Then
            range("A6").CurrentRegion.Select
            Selection.Subtotal GroupBy:=3, Function:=xlSum, TotalList:=Array(16), Replace:=False, PageBreaks:=False, _
                SummaryBelowData:=True
            Columns("Q:R").Select
            Selection.Delete Shift:=xlToLeft
        End If
    End If

    range("A6").CurrentRegion.Select
    Selection.ClearOutline
    

'DELETE GRAND TOTALS
Dim x As Long, lastrow As Long
lastrow = Cells(Rows.count, 16).End(xlUp).Row
For x = lastrow To 1 Step -1
    If Cells(x, 1).Value = "Grand Total" Or Cells(x, 2) = "Grand Total" Or Cells(x, 3) = "Grand Total" Then
        Rows(x).Delete
    End If
Next x


'MARK ZERO AS EXCLUDED
'lastrow = Cells(Rows.Count, 16).End(xlUp).Row
'For x = 7 To lastrow Step 1
'    If Cells(x, 16).Value = "0" Then
'        Cells(x, 16).Value = "excl."
'    End If
'Next x

pb.AddProgress 10

'CREATE AREA DIVISOR ON TOTALS
pb.AddCaption "Calculating area divisor on totals..."

If ThisWorkbook.Names("detail_prim_div").RefersToRange(1, 1).Value = "Yes" Then

primDivUnit = ThisWorkbook.Names("prim_div_unit").RefersToRange(1, 1).Value
primDivQty = ThisWorkbook.Names("prim_div_qty").RefersToRange(1, 1).Value

For x = 7 To lastrow Step 1
    If IsNumeric(Cells(x, 16).Value) = True And Cells(x, 1).Font.Bold = True Or _
    IsNumeric(Cells(x, 16).Value) = True And Cells(x, 2).Font.Bold = True Then
        Cells(x, 13).Value = Cells(x, 16).Value / primDivQty
        Cells(x, 13).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        Cells(x, 14).Value = "'/ " & primDivUnit
        Rows(x).RowHeight = 18
        Rows(x).Font.Bold = True
    End If
Next x
End If

pb.AddProgress 5

End Sub
