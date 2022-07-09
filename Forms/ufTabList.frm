VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufTabList 
   Caption         =   "Tab List"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "ufTabList.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "ufTabList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbCancel_Click()
' Purpose:
' Accepts:
' Returns:
    
    Unload Me
    
End Sub

Private Sub cbOK_Click()
' Purpose:
' Accepts:
' Returns:
    ' hide all except the selectio, if no selection no hidging

    
    Select Case tabList.MultiSelect
    Case fmMultiSelectSingle
        tabname = ufTabList.tabList.Value ' NOT A MULTIVALUE PROCESS
        If tabname <> "" Then
            Worksheets(tabname).Visible = True
            If cbActivate.Value = True Then Worksheets(tabname).Activate
            
            If HideAllOther.Value = True Then
                For Each w In Worksheets
                    If w.name <> tabname Then
                        w.Visible = False
                    End If
                Next
            End If
        End If
        
    Case fmMultiSelectMulti, fmMultiSelectExtended
        Dim oCol As Object
        Set oCol = CreateObject("scripting.dictionary")
        
        For el = 0 To tabList.ListCount - 1
            If tabList.Selected(el) = True Then
                tabname = tabList.List(el)
                Worksheets(tabname).Visible = True
                If cbActivate.Value = True Then Worksheets(tabname).Activate
                oCol.Item(tabname) = ""
            End If
        Next
        
        If HideAllOther.Value = True Then
            For Each w In Worksheets
                If Not oCol.exists(w.name) Then
                     w.Visible = False
                End If
            Next
        End If
        
        Set oCol = Nothing

    End Select
    
    Unload Me
    
End Sub

Private Sub UserForm_Activate()
' Purpose:
' Accepts:
' Returns:

    For Each w In Worksheets
        If w.Visible = False Then
            Me.tabList.AddItem w.name
        End If
    Next

End Sub
