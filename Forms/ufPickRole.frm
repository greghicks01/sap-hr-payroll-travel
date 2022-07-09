VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufPickRole 
   Caption         =   "Roles"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   OleObjectBlob   =   "ufPickRole.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufPickRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCancelled As Boolean

Private Sub cbCancel_Click()
' Purpose:
' Accepts:
' Returns:
    bCancelled = True
    Me.Hide
End Sub

Private Sub cbOK_Click()
' Purpose:
' Accepts:
' Returns:
    Me.Hide
End Sub

Private Sub TextBox1_Change()
' Purpose: allows the user to "locate" values quickly
' Accepts:
' Returns:

    For Each varItem In rolelist.List
        If InStr(1, varItem, UCase(TextBox1.Value)) > 0 Then
            rolelist.Text = varItem
            Exit For
        End If
    Next

End Sub

Private Sub UserForm_Initialize()
' Purpose: Reads all the data in the background sheet and adds to the list
' Accepts:
' Returns:

    iRow = 2
    While Sheets("Security Roles").Cells(iRow, "A") <> ""
        rolelist.AddItem Sheets("Security Roles").Cells(iRow, "A")
        iRow = iRow + 1
    Wend

End Sub

