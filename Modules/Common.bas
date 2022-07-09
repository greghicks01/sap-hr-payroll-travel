Attribute VB_Name = "Common"
Function locateHeader(src As Worksheet, colName As String) As Long
' Purpose: Locates the column "header" in row 1
' Accepts: src source worksheet, colName string in row 1 to locate
' Returns: column as an integer

    locateHeader = 1
    With src
        Do While .Cells(1, locateHeader) <> ""
            Select Case .Cells(1, locateHeader).Value
                Case Is = colName
                    Exit Function
            End Select
            locateHeader = locateHeader + 1
        Loop
        If .Cells(1, locateHeader) = "" Then locateHeader = 0
    End With
    
End Function

Function selectFileSystemItem(m As MsoFileDialogType, Optional fileFilter As String = "") As String
' Purpose:
' Accepts:
' Returns:

    Dim v As Variant
    
    With Application.FileDialog(m)
        .AllowMultiSelect = False
        If fileFilter <> "" Then
            
            .InitialFileName = fileFilter
        End If
        
        If .Show = -1 Then
            For Each v In .SelectedItems
                selectFileSystemItem = v
            Next
        Else
            selectFileSystemItem = "-1"
        End If
    End With
    
End Function

Sub removeFreezePanes()
' Purpose: removeFreezPanes Macro
' Accepts:
' Returns:
'
    For Each ws In ActiveWorkbook.Worksheets
        unhideAll
        hideSpecial
        If ws.Visible = True Then
            ws.FreezePanes = False
            ws.Range("b2").Select
            ws.FreezePanes = True
        End If
    Next
    
End Sub

Function ArrayToCollection(s() As Variant) As Collection
' Purpose: Coverts an Array to a collection object
' Accepts: s array as a variant
' Returns: loaded collection

    Dim c As Collection
    
    Set c = New Collection
    
    For Each x In s
        c.Add x
    Next
    
    Set ArrayToCollection = c
    
    Set c = Nothing
    
