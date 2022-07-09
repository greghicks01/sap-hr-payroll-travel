VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufInputValue 
   Caption         =   "-"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   OleObjectBlob   =   "ufInputValue.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufInputValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fpvCallBackFunction As String

Public Property Let callbackName(rhs As String)
' Purpose:
' Accepts:
' Returns:
    fpvCallBackFunction = rhs
End Property

Private Sub cmdbtnCancel_Click()
' Purpose:
' Accepts:
' Returns:
    Unload Me
End Sub

Private Sub cmdbutOK_Click()
' Purpose:
' Accepts:
' Returns:
    Me.Hide
    If fpvCallBackFunction <> "" Then
        CallByName cmdbutOK, fpvCallBackFunction, VbMethod, inputValue.Value
    Else
        MsgBox "Please correct the calling code", vbOKOnly, "Cannot execute function"
    End If
    Unload Me
End Sub
