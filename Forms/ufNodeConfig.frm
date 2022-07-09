VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufNodeConfig 
   Caption         =   "Add Org Stucture Node"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   OleObjectBlob   =   "ufNodeConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufNodeConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCancelled As Boolean

Private Sub CommandButton1_Click()
' Purpose:
' Accepts:
' Returns:
    Dim v As String
    v = ""
    If Not dispUFPickRole(v) Then
        Exit Sub
    End If
    
    If InStr(1, v, ";") > 0 Then
        arT = Split(v, ";")
    
        For Each sRole In arT
            roles.AddItem sRole
        Next
    Else
        roles.AddItem v
    End If
    
End Sub

Private Sub node_Change()
' Purpose:
' Accepts:
' Returns:
    
    If Me.node.Value = "org" Then
        Call taggedVisible("1", False)
    Else
        Call taggedVisible("1", True)
    End If
    
End Sub

Private Sub taggedVisible(sTag, bVis)
' Purpose:
' Accepts:
' Returns:

    For Each oControl In Me.Controls
        If oControl.Tag <> sTag Then
            oControl.Visible = bVis
        End If
    Next
    
End Sub

Private Sub cbOK_Click()
' Purpose:
' Accepts:
' Returns:
    bCancelled = False
    Me.Hide
End Sub

Private Sub cbCancel_Click()
' Purpose:
' Accepts:
' Returns:

    bCancelled = True
    Me.Hide
    
End Sub

Private Sub org_Change()
' Purpose: this value affects what we select in the next CB
' Accepts:
' Returns:
    Dim rDataRng As Range
    Dim wsData As Worksheet
    iTmp = 16 + Me.org.ListIndex
    Set wsData = Worksheets("Pay Scale Data")
    'sRange = Range(Cells(2, iTmp), Cells(13, iTmp)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    level.Clear
    
    'For Each sVal In wsData.Range(sRange)
        level.AddItem "DH" 'sVal
    'Next

End Sub

Private Sub sbFill_SpinDown()
' Purpose:
' Accepts:
' Returns:

    If fill.Value = 0 Then
        Exit Sub
    End If
    fill = fill - 1

End Sub

Private Sub sbFill_SpinUp()
' Purpose:
' Accepts:
' Returns:

    If CInt(fill.Value) < CInt(qty.Value) Then
        fill = fill + 1
    End If

End Sub

Private Sub sbQty_SpinDown()
' Purpose:
' Accepts:
' Returns:

    If CInt(qty) > 1 Then
    
        qty = qty - 1
        
        If CInt(qty) = 1 Then
            Me.position.Visible = True
        End If
        
    End If

End Sub

Private Sub sbQty_SpinUp()
' Purpose:
' Accepts:
' Returns:

    qty = qty + 1
    
    If qty > 1 Then
        Me.position.Visible = False
    End If

End Sub

Private Sub tbParentType_Change()
' Purpose:
' Accepts:
' Returns:

    Dim oControl As Control
    
    Me.node.Clear

    Me.node.AddItem ("pos")
    
    If Me.tbParentType.Value <> "org" Then
        Me.node.AddItem ("org")
    End If
    
End Sub


Private Sub UserForm_Initialize()
' Purpose:
' Accepts:
' Returns:
    Dim rDataRange As Range
    Dim cCell As Object
    
    Me.node.AddItem ("pos")
    Me.node.AddItem ("org")
    ' Data from Pay Scale Data
    rDataRng = Sheets("Pay Scale Data").Range("P1:R1")
    For Each sVal In rDataRng
        org.AddItem sVal
    Next
    
    roles.AddItem "ZGLOBAL_ORG"
    roles.AddItem "ZHR_EMPLOYEE"

End Sub
