Attribute VB_Name = "UserForms"

Function dispUF(ByRef v As String) As Boolean
' Purpose:
' Accepts:
' Returns:

    dispUF = True
    Load ufPayRollFilter
    ufPayRollFilter.Show
    btmp = ufPayRollFilter.cancelled
    iIdx = ufPayRollFilter.lbGroupID.ListIndex
    If iIdx >= 0 Then
        v = ufPayRollFilter.lbGroupID.List(iIdx)
    End If
    
    Unload ufPayRollFilter
    
    If btmp Or iIdx = -1 Then
        dispUF = False
    End If
    
End Function

Function dispUFExp(ByRef v As String) As Boolean
' Purpose:
' Accepts:
' Returns:
    Dim oControl As Control

    dispUFExp = True
    Load ufSelectExport
    ufSelectExport.Show
    
    btmp = ufSelectExport.bCancelled
    
    For Each oControl In ufSelectExport.Frame1.Controls
        If oControl.Value = True Then
            v = oControl.name
        End If
    Next
    
    Unload ufSelectExport
    
    If btmp Then
        dispUFExp = False
    End If
    
End Function

Function dispUFNC(ByRef v As String) As Boolean
' Purpose: Detects if user selected cancel or OK and returns T/F
' Accepts: v string for parent type
' Returns: Boolean for Cancel or OK and v with string data for later processing
    Dim oControl As Control

    dispUFNC = True
    Load ufNodeConfig
    ufNodeConfig.tbParentType = v
    v = ""
    ufNodeConfig.Show
    
    btmp = ufNodeConfig.bCancelled
    
    For Each oControl In ufNodeConfig.frData.Controls
        
        If oControl.Visible = True Or oControl.name = "position" Then
            v = v & oControl.name & ":"
            If TypeName(oControl) = "ListBox" Then
                
                For iList = 0 To oControl.ListCount - 1
                    v = v & oControl.List(iList) & ";"
                Next
                v = Left(v, Len(v) - 1) & "|"
            Else
                v = v & oControl.Value & "|"
            End If
        End If
    Next
    
    If v <> "" Then
        v = Left(v, Len(v) - 1)
    End If
    
    Unload ufNodeConfig
    
    If btmp Then
        dispUFNC = False
    End If
    
End Function

Function dispUFPickRole(ByRef v As String) As Boolean
' Purpose:
' Accepts:
' Returns:

    Dim oControl As Control

    dispUFPickRole = True
    Load ufPickRole
    v = ""
    ufPickRole.Show
    
    btmp = ufPickRole.bCancelled
    
    For el = 0 To ufPickRole.rolelist.ListCount - 1
        If ufPickRole.rolelist.Selected(el) = True Then
            v = v & ufPickRole.rolelist.List(el) & ";"
        End If
    Next
        
    On Error Resume Next
    v = Left(v, Len(v) - 1)
    On Error GoTo 0
    
    Unload ufPickRole
    
    If btmp Then
        dispUFPickRole = False
    End If

End Function

Function GetAColor() As Variant
' Purpose:
' Accepts:
' Returns:
'   Displays a dialog box and returns a
'   color value - or False if no color is selected
    Load ufColourPicker
    ufColourPicker.Show
    GetAColor = ColorValue
    Unload ufColourPicker
End Function
