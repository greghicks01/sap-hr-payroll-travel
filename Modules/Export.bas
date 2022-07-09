Attribute VB_Name = "Export"
Sub export_column_change()
' Purpose:
' Accepts:
' Returns:

    Dim srcWb As Workbook
    Dim dstWb As Workbook
    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim srcRow As Integer
    Dim dstRow As Integer
    
    unhideAll
    hideSpecial
    'copyACHireHeadings
    
    Set srcWb = Workbooks(1)
      
    Set dstWb = Workbooks.Add()
    For x = 2 To 3
        dstWb.Sheets(2).Delete
    Next
    
    iRow = 2
    While srcWb.Worksheets("Configuration").Cells(iRow, "A") <> ""
        If srcWb.Worksheets("Configuration").Cells(iRow, "A") = "TempPath" Then
            sTmpPath = srcWb.Worksheets("Configuration").Cells(iRow, "B")
        End If
        iRow = iRow + 1
    Wend
    dstWb.SaveAs sTmpPath & "temp"
    Set dstWs = dstWb.Sheets(1)
    
    srcRow = 1
    dstRow = 1
    
    For Each srcWs In srcWb.Sheets
        With srcWs
            If .Visible And locateHeader(srcWs, "exeID") <> 0 Then
                'Row Scan
                While .Cells(srcRow, locateHeader(srcWs, "Level")) <> "" ' Level is mandatory
                    'Col Scan
                    srcCol = 1
                    While .Cells(1, srcCol) <> "" ' All headers
                        dstWs.Cells(dstRow, srcCol) = .Cells(srcRow, srcCol)
                        If srcRow = 1 Then
                            dstWs.Cells(dstRow, srcCol).Interior.Color = .Cells(srcRow, srcCol).Interior.Color
                        End If
                        srcCol = srcCol + 1
                    Wend
                    srcRow = srcRow + 1
                    dstRow = dstRow + 1
                Wend
            End If
        End With
        srcRow = 2
    Next
    
    dstWb.Save

End Sub

Sub QTPTestPrep()
' Purpose:
' Accepts:
' Returns:
' P: Exports for use by QTP, organised by the "payroll" or identifier in user
' A:
' R:

    Dim srcW As Worksheet
    Dim dstW As Worksheet
    Dim srcWb As Workbook
    Dim dstWb As Workbook
    Dim srcRow As Integer
    Dim dstRow As Integer
    Dim payroll As String

    If Not dispUF(payroll) Then
        Exit Sub
    End If

    Set dstWb = Application.Workbooks.Open(SPath & "\CentrelinkSAPConsolRecords.xls")
    Set srcWb = Application.Workbooks("CentrelinkSAPConsolRecords.xlsm")
    Application.Windows.Arrange xlArrangeStyleHorizontal

    Call UnHideHidden(payroll)

    Set dstW = dstWb.Worksheets("Global")

    ' clears the data before copy out
    dstW.Activate
    Cells.Select
    Selection.Clear
    Range("A1").Select

    srcRow = 1
    dstRow = 1

    For Each srcW In srcWb.Worksheets
    
        With srcW
            If .Visible = True Then
                ' row scan
                While .Cells(srcRow, locateHeader(srcW, "Level")) <> ""
                    ' colScan
                    scrCol = 1
                    While .Cells(1, scrCol) <> ""
                        dstW.Cells(dstRow, scrCol) = srcW.Cells(srcRow, scrCol)
                        scrCol = scrCol + 1
                    Wend ' column scan
                    
                    srcRow = srcRow + 1
                    dstRow = dstRow + 1
                    
                Wend ' row scan
            
                srcRow = 2
            End If
            
        End With
        
    Next
    
    Application.DisplayAlerts = False
    
    dstWb.Close True, SPath & "\SAPConsolRecordsPay" & payroll & ".xls"
    Application.Windows.Arrange xlArrangeStyleHorizontal
    
    Application.DisplayAlerts = True

End Sub

Sub ExportPayroll()
' Purpose: Asks user to select a payroll area for export to top8 server
' Accepts: NIL
' Returns: NIL
    
    Dim srcWb       As Workbook, _
        dstWb       As Workbook
    
    Dim srcWs       As Worksheet, _
        dstWs       As Worksheet
    
    Dim srcRow      As Integer, _
        dstRow      As Integer, _
        dstWKS      As Integer
    
    Dim PathName    As String, _
        fileName    As String, _
        SheetName   As String, _
        payroll     As String, _
        sExportType As String
        
    If Not dispUF(payroll) Then
        Exit Sub
    End If
    
    If Not dispUFExp(sExportType) Then
        Exit Sub
    End If
           
    PathName = "S:\Automation\" '
    fileName = "SAP Consolidation Dataset ~ System ~ Client ~ Date ~.xlsx"
    SheetName = ""
    
    Call UnHideHidden(payroll)
      
    Set srcWb = Application.Workbooks("SAPConsolRecords.xlsm")
    Set dstWb = Application.Workbooks.Add
    Application.Windows.Arrange xlArrangeStyleHorizontal
    
    dstWKS = 1

    For Each srcWs In srcWb.Worksheets
        With srcWs
    
            If .Visible = True Then
                     
                srcRow = 1
                dstRow = 1
                
                'add extra pages
                If dstWKS > dstWb.Sheets.Count Then
                    dstWb.Sheets.Add After:=dstWb.Sheets(dstWb.Sheets.Count)
                End If
                
                Set dstWs = dstWb.Worksheets(dstWKS)
                dstWs.name = .name
                dstWs.Tab.Color = .Tab.Color
            
                If UCase(sExportType) = "ALL" Then
                    Call dataExport(srcWs, srcRow, dstWs, dstRow)
                Else
                    Call fnExpSpec(srcWs, srcRow, dstWs, dstRow)
                End If
                
                dstWKS = dstWKS + 1
                
                If SheetName = "" Then SheetName = .name
                
            End If
        End With
    Next
    ' final save
    
    srcRow = 2
    'Find first non-blank Activity row
    With srcWb.Worksheets(SheetName).Cells(srcRow, locateHeader(srcWb.Worksheets(SheetName), "SAPCLIENT"))
        System = UCase(Left(.Value, 3))
        Client = Right(.Value, 3)
     
        fileName = Replace(fileName, "~", payroll, 1, 1)                    ' Payroll 60
        fileName = Replace(fileName, "~", System, 1, 1)                     ' System R1D
        fileName = Replace(fileName, "~", Client, 1, 1)                     ' Client 222
        fileName = Replace(fileName, "~", Format(Now, "YYYY.MM.DD - HH.MM"), 1, 1)   ' Date
    End With
    
    dstWb.SaveAs fileName:=PathName & fileName, FileFormat:=xlWorkbookDefault
    dstWb.Close
    
    ' email me
    If sExportType <> "ALL" Then
        ' email to recipients
    End If
    ' send message to Tracy
End Sub

Sub dataExport(srcWs As Worksheet, srcRow As Integer, dstWs As Worksheet, dstRow As Integer)
' Purpose: Asks user to select a payroll area for export to top8 server
' Accepts: NIL
' Returns: NIL

    'srcRefCol = locateLevelHeader(srcWs)
    
    With srcWs
        While .Cells(srcRow, locateHeader(srcWs, "Level")) <> "" ' Level is mandatory
            'Col Scan
            srcCol = 1
            While .Cells(1, srcCol) <> "" ' All headers
                If dstRow = 1 Then
                    dstWs.Cells(dstRow, srcCol).Interior.Color = .Cells(srcRow, srcCol).Interior.Color
                End If
                dstWs.Cells(dstRow, srcCol).Value = .Cells(srcRow, srcCol).Value
                srcCol = srcCol + 1
            Wend
            srcRow = srcRow + 1
            dstRow = dstRow + 1
        Wend
    End With
    dstRow = 2
    
End Sub

Sub fnExpSpec(srcWs As Worksheet, srcRow As Integer, dstWs As Worksheet, dstRow As Integer)
' Purpose: Exports a dataset in test users format.
' Accepts: NIL
' Returns: NIL

    'A       B       C           D           E               F               G       H           I
    'Parent  Payroll Pers_Area   Pers_Sub    Org_Unit_Name   Org_Unit_No.    AGS_Nos Position    Logon_Id
    ' Output row 1 cells
    Set cHeadCells = CreateObject("scripting.dictionary")
    cHeadCells.Item("E1") = ""
    cHeadCells.Item("H1") = ""
    cHeadCells.Item("J1") = ""
        
    'output row 2 cells
    Set cHead2Cells = CreateObject("scripting.dictionary")
    cHead2Cells.Item("A2") = "Org_Unit_Name"
    cHead2Cells.Item("B2") = "Org_Unit_No."
    cHead2Cells.Item("C2") = "AGS_Nos"
    cHead2Cells.Item("D2") = "Position"
    cHead2Cells.Item("E2") = "Logon_Id"
    cHead2Cells.Item("F2") = "Last_Name"
    cHead2Cells.Item("G2") = "First_Name"
    cHead2Cells.Item("H2") = "Level" '"Pref_Name"
    cHead2Cells.Item("I2") = "Sup_pos_no." '"Level"
    cHead2Cells.Item("J2") = "DT_PP13_Roles" '"Sup_pos_no."
    cHead2Cells.Item("K2") = "Gender" '"Activity_Group"

    With srcWs
    
        While .Cells(srcRow, locateHeader(srcWs, "Level")) <> "" ' Level is mandatory
            'Col Scan
            If dstRow = 1 Then
                'Prep header
                
                For Each sKey In cHeadCells
         
                    Select Case sKey
                    'Case "C1"
                        'dstWs.Range(sKey).Value = .Cells(2, locateHeader(srcWs, "Parent"))
                    Case "E1"
                        dstWs.Range("A1") = "Payroll Area = " & .Cells(2, locateHeader(srcWs, "Payroll"))
                    Case "H1"
                        dstWs.Range("C1") = "Pers Area = " & .Cells(2, locateHeader(srcWs, "Pers_Area"))
                    Case "J1"
                        dstWs.Range("E1") = "Pers Sub Area = " & .Cells(2, locateHeader(srcWs, "Pers_Sub"))
                    End Select
                Next
                
                For Each sKey In cHead2Cells
                    dstWs.Range(sKey).Value = cHead2Cells.Item(sKey)
                Next
                
                dstRow = 2

            Else
            
                dstCol = 1
                For Each sKey In cHead2Cells ' All headers
                    temp = cHead2Cells.Item(sKey)
                    tempV = ""
                    If sKey = "K2" Then
                        If .Cells(srcRow, locateHeader(srcWs, CStr(temp))).Value <> "" Then
                            If InStr(1, .Cells(srcRow, locateHeader(srcWs, CStr(temp))).Value, "~") <> 0 Then
                                tempA = Split(.Cells(srcRow, locateHeader(srcWs, CStr(temp))), "~")
                                tempV = tempA(1)
                            Else
                                tempV = .Cells(srcRow, locateHeader(srcWs, cHead2Cells.Item(sKey)))
                            End If
                        End If
                    Else
                        tempV = .Cells(srcRow, locateHeader(srcWs, CStr(temp)))
                    End If
                    
                    dstWs.Cells(dstRow, dstCol) = tempV
                    
                    If .Cells(srcRow, locateHeader(srcWs, "Done")).Value = "F" Then
                        sRange = Range(Cells(srcRow, 1), Cells(srcRow, 12)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                        dstWs.Range(sRange).Interior.Color = RGB(255, 0, 0)
                    End If
                                                         
                    dstCol = dstCol + 1
                Next
            End If
        
            srcRow = srcRow + 1
            dstRow = dstRow + 1
            
        Wend
        
        'Adjust the titles, widths and cell borders
        iColCount = 1
        With srcWs.Parent
            Set sTmp = .Sheets("User Headers")
            While sTmp.Cells(1, iColCount) <> ""
                dstWs.Cells(2, iColCount) = sTmp.Cells(1, iColCount)
                With dstWs.Cells(2, iColCount).Font
                    .name = sTmp.Cells(1, iColCount).Font.name
                    .Bold = sTmp.Cells(1, iColCount).Font.Bold
                    .Size = sTmp.Cells(1, iColCount).Font.Size
                    .Strikethrough = sTmp.Cells(1, iColCount).Font.Strikethrough
                    .Superscript = sTmp.Cells(1, iColCount).Font.Superscript
                    .Subscript = sTmp.Cells(1, iColCount).Font.Subscript
                    .OutlineFont = sTmp.Cells(1, iColCount).Font.OutlineFont
                    .Shadow = sTmp.Cells(1, iColCount).Font.Shadow
                    .Underline = sTmp.Cells(1, iColCount).Font.Underline
                    .ThemeColor = sTmp.Cells(1, iColCount).Font.ThemeColor
                    .TintAndShade = sTmp.Cells(1, iColCount).Font.TintAndShade
                End With
                
                dstWs.Cells(2, iColCount).ColumnWidth = sTmp.Cells(1, iColCount).ColumnWidth
                
                iColCount = iColCount + 1
            Wend
        End With
        
        Range("A3", Cells(dstRow - 1, "K")).Select

        Call selectionProperties(xlEdgeTop)
        Call selectionProperties(xlEdgeBottom)
        Call selectionProperties(xlEdgeLeft)
        Call selectionProperties(xlInsideVertical)
        Call selectionProperties(xlEdgeRight)
         
        Range("A2", Cells(2, "K")).Select
       
        Call selectionProperties(xlEdgeTop)
        Call selectionProperties(xlEdgeBottom)
        Call selectionProperties(xlEdgeLeft)
        Call selectionProperties(xlInsideVertical)
        Call selectionProperties(xlEdgeRight)
        
        With Selection
            .HorizontalAlignment = xlGeneral
            .VerticalAlignment = xlBottom
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        
        Range("b3").Select
        ActiveWindow.FreezePanes = True
        
        'Change Name of the Tag?
 
    End With
    
    dstRow = 2
    
End Sub

Sub selectionProperties(BorderIndexVal As XlBordersIndex)

        With Selection.Borders(BorderIndexVal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlMedium
        End With
    
End Sub


