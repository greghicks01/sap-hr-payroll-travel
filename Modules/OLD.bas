Attribute VB_Name = "OLD"
Sub getColourIndex()
' Purpose:
' Accepts:
' Returns:

 For Each ws In Worksheets
        With ws
            If .Visible = True Then
            
                   Debug.Print .Tab.ColorIndex
            End If
        End With
    Next

End Sub

Sub Macro4()
' Purpose:
' Accepts:
' Returns:
'
' Macro5 Macro
'
Dim s As Worksheet
'
    For Each s In Application.ActiveWorkbook.Worksheets
    
        With s
    
            If s.Visible = True Then
                .Activate
                .Range("A1:AH300").AutoFilter field:=13, Criteria1:="APS6"
            End If
        End With
        
    Next

End Sub

Sub Macro6()
' Purpose:
' Accepts:
' Returns:
'
' Macro6 Macro
'

'
For Each s In Application.ActiveWorkbook.Worksheets

    With s

        If s.Visible = True Then
            .Activate
            Range("G2").Select
            ActiveWindow.SmallScroll Down:=300
            Range("G2:G199").Select
            Selection.ClearContents
        End If
    End With
    
Next
    
End Sub
Sub addPassword()
' Purpose:
' Accepts:
' Returns:

    For Each ws In Worksheets
        With ws
            If .Visible = True Then
                destRow = 2
                
                While .Cells(destRow, locateHeader(ws, "Level")).Value <> ""
                
                    If .Cells(destRow, locateHeader(ws, "Activity_Group")).Value <> "" Then
                        'LEAVE
                        Select Case .Tab.ColorIndex
                            Case Is = 49 ' Blue CL
                                rl = "RL"
                                pl = "PM"
                            Case Is = 55 ' Green MC
                                rl = "RF"
                                pl = "PF"
                            Case Is = 10 ' Red HS
                                rl = "RL"
                                pl = "PM"
                        End Select
                        .Cells(destRow, locateHeader(ws, "REC_Leave")).Value = rl
                        .Cells(destRow, locateHeader(ws, "Long_Service_FT")).Value = "LS"
                        .Cells(destRow, locateHeader(ws, "Long_Service_PT")).Value = "LP"
                        .Cells(destRow, locateHeader(ws, "Per_Leave")).Value = pl
                        
                        'password
                        .Cells(destRow, locateHeader(ws, "Password")).Value = "welcome"
                    End If
                
                    destRow = destRow + 1
                    
                Wend
            End If
        End With
        
    Next
    
End Sub


Sub correctActivityCodes()
' Purpose:
' Accepts:
' Returns:

    Dim srcWb As Workbook
    Dim srcW As Worksheet
    Dim srcRow As Integer
    Dim column As Integer
    
    column = 15

    Set srcWb = Application.Workbooks("SAPConsolRecords.xlsm")

    For Each srcW In srcWb.Worksheets
    
        srcRow = 2

        With srcW

            If .Visible = True Then
            
                'row scan
                While .Cells(srcRow, 1) <> ""
                    With .Cells(srcRow, column)
                        If Right(.Value, 1) = ";" Then
                            .Value = Mid(.Value, 1, Len(.Value) - 1)
                        End If
                        If .Value <> "" And InStr(.Value, "~") = 0 Then
                            .Value = "r1dclnt222~" & .Value
                        End If
                    End With
                    
                    srcRow = srcRow + 1

                Wend
            End If
        End With
    Next
    
End Sub

Sub copySetPages()
' Purpose:
' Accepts:
' Returns:

    Dim destW As Worksheet
    Dim srcW As Worksheet
    Dim dwb As Workbook
    Dim swb As Workbook
    Dim dRow As Integer
    Dim sRow As Integer
    Const dCompCol = 13
    Const dDataCol = 8
    Const sCompCol = 4
    Const sDataCol = 3
      
    Set dwb = Application.Workbooks("SAPConsolRecords.xlsm")
    Set swb = Application.Workbooks("positionAPS4-5Creation.xls")
    Set srcW = swb.ActiveSheet
    'Application.Windows.Arrange xlArrangeStyleVertical
    
    dRow = 1
    sRow = 2

    For Each destW In dwb.Worksheets
    
        dRow = 2
    
        With destW
        
            If .Visible = True Then
        
                While .Cells(dRow, 1) <> ""
                    
                    dRow = dRow + 1
                    
                    'match the source D to dest M
                    If .Cells(dRow, dCompCol).Value = srcW.Cells(sRow, sCompCol).Value Then
                    
                        .Cells(dRow, dDataCol).Value = srcW.Cells(sRow, sDataCol).Value
                       
                        sRow = sRow + 1
                        
                        If srcW.Cells(sRow, sDataCol).Value = "" Then Exit Sub
                        
                    End If
                    
                Wend
            
            End If
            
        End With
        
    Next

End Sub

Sub copyPerSubtoDefault()
' Purpose:
' Accepts:
' Returns:

    Dim srcWb As Workbook
    Dim dstWb As Workbook
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim srcRow As Integer
    Dim dstRow As Integer
    
    Set srcWb = Application.Workbooks("SAPConsolRecords.xlsm")
    Set dstWb = Application.Workbooks("Default.xls")
    
    Set dstWs = dstWb.Worksheets("Global")
    dstRow = 2
    '
    On Error Resume Next
    
    For Each srcWs In srcWb.Worksheets
        
        With srcWs
        
            If .Visible = xlSheetVisible Then
        
                srcRow = 2
                
                ' Parent
                While .Cells(srcRow, 1) <> ""
                
                    ' AGS Number
                    While .Cells(srcRow, 7) = "" And .Cells(srcRow, 1) <> ""
                        srcRow = srcRow + 1
                    Wend
                    
                    'copy pers area to default
                    If .Cells(srcRow, 1) <> "" Then
                    
                        dstWs.Cells(dstRow, 4).Value = .Cells(srcRow, 3).Value
                        dstWs.Cells(dstRow, 5).Value = .Cells(srcRow, 4).Value
                        
                        dstRow = dstRow + 1
                        srcRow = srcRow + 1
                        
                    End If
                    
                Wend
            
            End If
            
        End With
    Next
    
End Sub

Sub copyAll()
' Purpose:
' Accepts:
' Returns:

    Dim s As Worksheet
    Dim awb As Workbook
    Dim nwb As Workbook
    Dim nrow As Integer
    
    nrow = 1
    
    Set awb = Application.ActiveWorkbook
    Set nwb = Application.Workbooks.Add
    Application.Windows.Arrange xlArrangeStyleHorizontal

    For Each s In awb.Worksheets
    
        With s

            If .Visible = True Then

                awb.Activate
                .Activate

                If .name = "ACHire" Then
                    Range(CStr(nrow) & ":" & CStr(nrow)).Copy
                    nwb.Activate
                    ActiveSheet.Paste
                End If

                awb.Activate
                
                t = Application.WorksheetFunction.CountA(s.Range("B:B")) - 1

                For x = 2 To Application.WorksheetFunction.CountA(s.Range("B:B")) - 1
                
                    If UCase(Trim(ActiveSheet.Cells(x, 13).Value)) = "APS5" Or UCase(Trim(ActiveSheet.Cells(x, 13).Value)) = "APS4" Then
                        nrow = nrow + 1
                        ActiveSheet.Range(CStr(x) & ":" & CStr(x)).Copy
                        nwb.Activate
                        ActiveSheet.Cells(CStr(nrow), 1).Select
                        ActiveSheet.Paste
                        awb.Activate
                    End If
                Next

            End If

        End With

    Next

End Sub

Sub copyPositionsFrom()
' Purpose:
' Accepts:
' Returns:

    Dim destW As Worksheet
    Dim srcW As Worksheet
    Dim dwb As Workbook
    Dim swb As Workbook
    Dim dRow As Integer
    Dim sRow As Integer
    Const dCompCol = 13
    Const dDataCol = 8
    Const sCompCol = 4
    Const sDataCol = 3
      
    Set dwb = Application.Workbooks("SAPConsolRecords.xlsm")
    Set swb = Application.Workbooks("positionAPS4-5Creation.xls")
    Set srcW = swb.ActiveSheet
    'Application.Windows.Arrange xlArrangeStyleVertical
    
    dRow = 1
    sRow = 2

    For Each destW In dwb.Worksheets
    
        dRow = 2
    
        With destW
        
            If .Visible = True Then
        
                While .Cells(dRow, 1) <> ""
                    
                    dRow = dRow + 1
                    
                    'match the source D to dest M
                    If .Cells(dRow, dCompCol).Value = srcW.Cells(sRow, sCompCol).Value Then
                    
                        .Cells(dRow, dDataCol).Value = srcW.Cells(sRow, sDataCol).Value
                       
                        sRow = sRow + 1
                        
                        If srcW.Cells(sRow, sDataCol).Value = "" Then Exit Sub
                        
                    End If
                    
                Wend
            
            End If
            
        End With
        
    Next

End Sub

Sub copyStaffOut()
' Purpose:
' Accepts:
' Returns:
    
    Dim srcWb As Workbook
    Dim dstWb As Workbook
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim srcRow As Integer
    Dim dstRow As Integer
    Dim srcColArr() As String
    Dim dstColArr() As String
    Dim srcColCln As Collection
    Dim dstColCln As Collection
    
    Const dCompCol = 13
    Const sCompCol = 4
    Const dDataCol = 8
    Const sDataCol = 3
     
    srcColArr = Split("7 8 9 10 11 12 13 14 15 16 17 20 21 23 25 26 27 28 29 30 31 32 33 34")
    dstColArr = Split("1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24")
       
    Set srcColCln = ArrayToCollection(srcColArr)
    Set dstColCln = ArrayToCollection(dstColArr)
    
    dstRow = 2
    
    Set srcWb = Application.Workbooks("CentrelinkSAPConsolRecords.xlsm")
    Set dstWb = Application.Workbooks.Add
    Set dstWs = dstWb.Worksheets("Sheet1")
    
    Application.Windows.Arrange xlArrangeStyleHorizontal

    For Each srcWs In srcWb.Worksheets
    
        srcRow = 2
          
        With srcWs
        
            If .Visible = xlSheetVisible Then
                
                If .name = "ACHire" Then
                    
                    For x = 1 To srcColCln.Count
                    
                        dstWs.Cells(1, CInt(dstColCln.Item(x))).Value = .Cells(1, CInt(srcColCln.Item(x))).Value
                        
                    Next
                    
                End If
        
                While .Cells(srcRow, 1) <> ""
                
                    
                    If .Cells(srcRow, 7) <> "" Then
                    
                        For x = 1 To srcColCln.Count
                    
                            dstWs.Cells(dstRow, CInt(dstColCln(x))).Value = .Cells(srcRow, CInt(srcColCln(x))).Value
                            
                        Next
                        
                        dstRow = dstRow + 1
                        
                    End If
                        
                    srcRow = srcRow + 1
                
                Wend
                
            End If
                
        End With
    
    Next
    
End Sub

