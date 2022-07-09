Attribute VB_Name = "Import"

Sub import_column_change()
' Purpose:
' Accepts:
' Returns:

    Dim srcWs As Worksheet
    Dim dstWs As Worksheet
    Dim srcWb As Workbook
    Dim dstWb As Workbook
    Dim srcRow As Integer
    Dim dstRow As Integer
    
    Set dstWb = Workbooks("CentrelinkSAPConsolRecords.xlsm")
    Set srcWb = Workbooks("temp.xlsx")
    Set srcWs = srcWb.Sheets(1)
    
    srcRow = 1
    dstRow = 1
    
    For Each dstWs In dstWb.Sheets
        import dstWs, srcWs, srcRow, dstRow
    Next
    
    srcWb.Close False
    
    dstWb.Activate
    
    copyACHireHeadings
    
End Sub

Function import(ByRef dstWs As Worksheet, ByRef srcWs As Worksheet, ByRef srcRow As Integer, ByRef dstRow As Integer)
' Purpose:
' Accepts:
' Returns:

    dstRefCol = locateHeader(dstWs, "Level")
    srcRefCol = locateHeader(srcWs, "Level")
    
    With dstWs
        If .Visible Then
            'Row Scan
            While srcWs.Cells(srcRow, srcRefCol) <> "" And .Cells(dstRow, dstRefCol) <> "" ' Level is mandatory
                'Col Scan
                srcCol = 1
                While srcWs.Cells(1, srcCol) <> "" ' All headers
                    .Cells(dstRow, srcCol) = srcWs.Cells(srcRow, srcCol)
                    If dstRow = 1 Then
                        .Cells(dstRow, srcCol).Interior.Color = srcWs.Cells(srcRow, srcCol).Interior.Color
                    End If
                    srcCol = srcCol + 1
                Wend
                srcRow = srcRow + 1
                dstRow = dstRow + 1
            Wend
        End If
    End With
    dstRow = 2
    
End Function

Sub copyResultsFromDefault()
' Purpose:Reads in the results from a QTP execution
' Accepts: NIL
' Returns:
    
    Dim srcW As Worksheet
    Dim dstW As Worksheet
    Dim srcWb As Workbook
    Dim dstWb As Workbook
    Dim srcRow As Integer
    Dim dstRow As Integer
    
    Set dstWb = Application.Workbooks("SAPConsolRecords.xlsm")
    temp = selectFileSystemItem(msoFileDialogFolderPicker)
    If temp = -1 Then Exit Sub
    
    ' auto drill
    ' locate the test executed and find the lastest results?
    
    Set srcWb = Application.Workbooks.Open(temp & "\Report\Default.xls") '("S:\Automation\SAPQTP\Control Record Scripts from Mark\SAPConsolRecords.xls") '
    Application.Windows.Arrange xlArrangeStyleHorizontal
    
    Application.ScreenUpdating = False
    
    Set srcW = srcWb.Worksheets("Global")
    Call UnHideHidden(srcW.Cells(2, locateHeader(srcW, "exeID")))
    hideSpecial
    
    srcRow = 2
    
    srcRefCol = locateHeader(srcW, "Level")

    For Each dstW In dstWb.Worksheets
        dstRefCol = locateHeader(dstW, "Level")
        dstRow = 2

        With dstW
            If .Visible = True Then
                'row scan
                While srcW.Cells(srcRow, srcRefCol) <> "" And .Cells(dstRow, dstRefCol) <> ""
                    'column scan
                    ColCount = 1
                    While .Cells(1, ColCount) <> ""
                        With .Cells(dstRow, ColCount)
                            If .Value <> srcW.Cells(srcRow, ColCount).Value Then
                                .Value = srcW.Cells(srcRow, ColCount)
                            End If
                        End With
                        ColCount = ColCount + 1
                    Wend
                    srcRow = srcRow + 1
                    dstRow = dstRow + 1
                Wend
            End If
        End With
    Next
    
    Application.ScreenUpdating = True
    
    srcWb.Close True
    
    dstWb.Activate
    
    ActiveWindow.WindowState = xlMaximized

End Sub
