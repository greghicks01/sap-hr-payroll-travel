Attribute VB_Name = "Update"
'Sub prepareCleanPayroll()
'' Purpose:
'' Accepts:
'' Returns:
'    Dim cols As Variant
'    Dim w As Worksheet
'    Dim srcCol As Variant
'    Dim srcRow As Long
'
'    strDlgMsg = "This will clear a number of columns in the nominated payroll" & vbCrLf & _
'              "ensure you really want to do this"
'
'
'    If MsgBox(strDlgMsg, vbOKCancel, "Confirm prepare") = vbCancel Then Exit Sub
'
'    columnsStr = "1,5,6,8,14,17,18"
'    cols = Split(columnsStr, ",")
'
'    For Each w In Worksheets
'        With w
'            If .Visible = xlSheetVisible Then
'                srcRow = 2
'                While .Cells(srcRow, 13) <> ""
'                    For Each srcCol In cols
'                        .Cells(srcRow, CLng(srcCol)).Value = ""
'                    Next
'                    srcRow = srcRow + 1
'                Wend
'            End If
'        End With
'    Next
'
'End Sub

'Sub updateActivitySystem(newSystem As String)
'' Purpose:
'' Accepts:
'' Returns:
'
'    Dim w As Worksheet
'    Dim temp As Variant
'    Const delimiter = "~"
'
'    If newSystem = "" Then Exit Sub
'
'    For Each w In Worksheets
'        With w
'            If .Visible = xlSheetVisible Then
'                While .Cells(srcRow, 13) <> ""
'                    If .Cells(srcRow, 15) <> "" Then
'                        temp = Split(.Cells(srcRow, columnID).Value, delimiter)
'                        .Cells(srcRow, columnID).Value = newSystem & delimiter & temp(1)
'                    End If
'                Wend
'            End If
'        End With
'    Next
'
'End Sub

'Sub PAYAREA()'
'    Dim src As Worksheet
'    Set src = ActiveSheet
'    curRow = 2
'
'    With src
'        While .Cells(curRow, locateHeader(src, "Pers_Area")) <> ""
'            Select Case Left(.Cells(curRow, locateHeader(src, "Pers_Area")).Value, 1)
'            Case "C"
'                ps_area = "CL"
'            Case "H"
'                ps_area = "HS"
'            Case "M"
'                ps_area = "MC"
'            End Select
'
'            .Cells(curRow, locateHeader(src, "PS_Area")).Value = ps_area
'            curRow = curRow + 1
'        Wend
'    End With
'
'End Sub

Sub partTimeUpdate()
' Purpose:
' Accepts:
' Returns:

    Dim w As Worksheet

    For Each w In Worksheets
        If w.Visible = True And locateHeader(w, "exeID") <> 0 Then
            iRow = 1
            With w
                While .Cells(iRow, locateHeader(w, "Level")) <> ""
                    If .Cells(iRow, locateHeader(w, "Employee_Group")) = "B" Then
                        .Cells(iRow, locateHeader(w, "PA40_i0220_Superannuation_Status")) = "PH"
                        .Cells(iRow, locateHeader(w, "PA40_i0007_PartTime_Schedule")) = "001 0017"
                    End If
                    iRow = iRow + 1
                Wend
            End With
        End If
    Next

End Sub

Sub cleanActivity()
'p removes leading sections of activity (clnt~)

    Dim w As Worksheet
    For Each w In ActiveWorkbook.Sheets
        If w.Visible Then
            iRow = 2
            With w
            Do While .Cells(iRow, locateHeader(w, "Level")) <> ""
                If InStr(1, .Cells(iRow, locateHeader(w, "Activity_Group")), "~") <> 0 Then
                    tmp = Split(.Cells(iRow, locateHeader(w, "Activity_Group")), "~")
                    .Cells(iRow, locateHeader(w, "Activity_Group")) = tmp(1)
                End If
                iRow = iRow + 1
            Loop
            End With
            
        End If
    Next
End Sub
