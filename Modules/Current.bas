Attribute VB_Name = "Current"

Sub RemoveFilters()
Attribute RemoveFilters.VB_ProcData.VB_Invoke_Func = " \n14"
' Purpose:REmoveFilters Macro
' Accepts:
' Returns:
    Dim s As Object
    '
    On Error Resume Next
    
    For Each s In Application.ActiveWorkbook.Worksheets
        
        With s
            If .Visible = True Then
                .Activate
                .ShowAllData
            End If
        End With
    Next
    
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
' Purpose: tidies up the sheets formats
' Accepts:
' Returns
    Dim s As Object
    
    For Each s In Application.ActiveWorkbook.Worksheets
    
        If s.Visible = True Then
            s.Activate
            ActiveWindow.Zoom = 100
            Cells.Select
            With Selection.Font
                .name = "Calibri"
                .Size = 9
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With Selection.Font
                .name = "Calibri"
                .Size = 9
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With Selection.Font
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
            End With
            With Selection
                .Rows.AutoFit
                .Columns.AutoFit
            End With
        End If
        
    Next
End Sub
End Function

Sub UnHideHidden(Optional ByVal payroll As String = "60")
' Purpose: Unhides all the sheets that have the header with exeID and matching identifier
' Accepts: DataSet Identifier
' Returns:

    Dim srcW As Worksheet

    Set srcWb = Application.Workbooks("SAPConsolRecords.xlsm")

    For Each srcW In srcWb.Worksheets

        With srcW
    
            tmp = locateHeader(srcW, "exeID")
            If tmp = 0 Then .Visible = False
            
            If tmp <> 0 Then
                If UCase(.Cells(2, tmp)) = UCase(payroll) Then 'B   Payroll
                    .Visible = True
                End If
            End If
        End With
        
    Next
    
    ' work around in case where all sheets are hidden before we unhide one.
    For Each srcW In srcWb.Worksheets
    
        tmp = locateHeader(srcW, "exeID")

        With srcW
            If tmp <> 0 Then
                If UCase(.Cells(2, tmp)) <> UCase(payroll) Then 'B   Payroll
                    .Visible = xlSheetHidden
                End If
            Else
                .Visible = xlSheetHidden
            End If
        End With
    Next

End Sub

Sub unhideAll()
' Purpose: unhides the lot
' Accepts:
' Returns:

    On Error Resume Next
    Set srcWb = Application.Workbooks("SAPConsolRecords.xlsm")

    For Each srcW In srcWb.Worksheets
        With srcW
            .Visible = True
        End With
    Next
        
End Sub

Sub hideSpecial()
' Purpose: hides sheets not involved in data creation
' Accepts:
' Returns:
   Dim s As Worksheet
   
    Set srcWb = Application.Workbooks("SAPConsolRecords.xlsm")
   
    For Each s In srcWb.Worksheets
        If locateHeader(s, "exeID") = 0 Then
            s.Visible = False
        End If
    Next

End Sub

' TODO: Improve to work from the ActiveSheet
Sub copyACHireHeadings()
' P: Copies headings from AC to all other useable pages
' A: Nil
' R: Nil
    Dim dstW As Worksheet, srcWb As Workbook, srcWs As Worksheet
    
    On Error Resume Next
    Set srcWb = Application.Workbooks("SAPConsolRecords.xlsm")
    Set srcWs = srcWb.Worksheets("ACHire")
        
    Application.ScreenUpdating = False
    
    For Each dstW In srcWb.Worksheets
        colCnt = 1
        With dstW
            .AutoFilterMode = False ' Turn off filter

            If .name <> srcWs.name And .Visible = True Then
                'col scan
                While srcWs.Cells(1, colCnt) <> ""
                    .Cells(1, colCnt).Value = srcWs.Cells(1, colCnt).Value
                    .Cells(1, colCnt).Interior.Color = srcWs.Cells(1, colCnt).Interior.Color
                    .Cells(1, colCnt).Columns.AutoFit
                    colCnt = colCnt + 1
                Wend
            End If
            .Range("A1").AutoFilter ' turn on filter
        End With
    Next
    
    Application.ScreenUpdating = True
    
End Sub

Sub scanAndCleanRoles()
' Purpose: Scans each sheet for findReplace on Roles
' Accepts:
' Returns:

    Dim s As Worksheet

    For Each s In Application.ActiveWorkbook.Worksheets
    
        srcRow = 2
        
        With s
            If .Visible = xlSheetVisible Then
        
                While .Cells(srcRow, locateHeader(s, "Level")) <> ""
                    If .Cells(srcRow, locateHeader(s, "Activity_Group")) <> "" Then
                        totReplace = totReplace + findReplace(.Cells(srcRow, locateHeader(s, "Activity_Group"))) 'O   Activity_Group
                    End If
                    While InStr(1, .Cells(srcRow, locateHeader(s, "Activity_Group")), ";;") <> 0 'O   Activity_Group
                        .Cells(srcRow, locateHeader(s, "Activity_Group")) = Replace(.Cells(srcRow, locateHeader(s, "Activity_Group")), ";;", ";") 'O   Activity_Group
                    Wend
                    srcRow = srcRow + 1
                Wend
            End If
        End With
    Next
    Debug.Print totReplace

End Sub

Private Function findReplace(c As Range) As Integer
' Purpose: performs find replace on data defined by c
' Accepts: Range to look in
' Returns:
    '
    srcRow = 2
    With Worksheets("Role Translation")
        While .Cells(srcRow, 1) <> "" '

            findVal = .Cells(srcRow, 1)
            tempBefore = c.Value
            tempAfter = Replace(tempBefore, .Cells(srcRow, 1), .Cells(srcRow, 2))
            If tempBefore <> tempAfter Then
                c.Value = Replace(tempAfter, ";;", ";")
                findReplace = findReplace + 1
            End If
            srcRow = srcRow + 1
        Wend
    End With
  
End Function


Private Sub dataTransfer(srcWs As Worksheet, ByRef srcRow As Integer, dstWs As Worksheet, ByRef dstRow As Integer)
' Purpose: Copies data from source to destintation sheets
' Accepts: NIL
' Returns: NIL
    Dim srcCol As Integer

    With srcWs
        ' row scan
        While .Cells(srcRow, locateHeader(srcWs, "Level")) <> "" 'M   Level

            ' colScan
            scrCol = 1
            While .Cells(1, scrCol) <> ""
                dstWs.Cells(dstRow, scrCol) = srcWs.Cells(srcRow, scrCol)
                scrCol = scrCol + 1
            Wend ' column scan
            
            srcRow = srcRow + 1
            dstRow = dstRow + 1
            
        Wend ' row scan
    End With
    
End Sub
