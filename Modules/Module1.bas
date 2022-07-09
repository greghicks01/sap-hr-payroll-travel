Attribute VB_Name = "Module1"

Sub greatestAGSIssued()
' Purpose:
' Accepts:
' Returns:
Dim ws As Worksheet
Dim biggest As Long
    unhideAll
    hideSpecial
    For Each ws In Worksheets
        With ws
            If .Visible Then
                srcRow = 2
                If locateHeader(ws, "exeID") <> 0 Then
                    While .Cells(srcRow, locateHeader(ws, "Level")) <> ""
                        If .Cells(srcRow, locateHeader(ws, "AGS_Nos")) <> "" Then
                            If .Cells(srcRow, locateHeader(ws, "AGS_Nos")) > biggest Then
                                biggest = .Cells(srcRow, locateHeader(ws, "AGS_Nos"))
                            End If
                        End If
                    srcRow = srcRow + 1
                    Wend
                End If
            End If
        End With
    Next
    Debug.Print biggest
End Sub
