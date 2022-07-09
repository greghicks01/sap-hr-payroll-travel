VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPayRollFilter 
   Caption         =   "Group Execution Identifiers"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3885
   OleObjectBlob   =   "ufPayRollFilter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPayRollFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public cancelled As Boolean

Private Sub cbCancel_Click()
' Purpose:
' Accepts:
' Returns:
    Me.Hide
    cancelled = True
End Sub

Private Sub cbOK_Click()
' Purpose:
' Accepts:
' Returns:
    cancelled = False
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
' Purpose:
' Accepts:
' Returns:

    'read off the top row where ColumnHead = "exeID
    cancelled = False
    Dim ws As Worksheet
    
    lbGroupID.Clear
    
    For Each ws In Workbooks(1).Worksheets
        tmp = locateHeader(ws, "exeID")
        bAlreadyIn = False
        If tmp <> 0 Then
            tmp = ws.Cells(2, tmp).Value
            For t = 0 To Me.lbGroupID.ListCount - 1
                B = lbGroupID.List(t)
                If CStr(B) = tmp Then
                    bAlreadyIn = True
                    Exit For
                End If
            Next
            If Not bAlreadyIn Then
                lbGroupID.AddItem tmp
            End If
        End If
    Next
End Sub
