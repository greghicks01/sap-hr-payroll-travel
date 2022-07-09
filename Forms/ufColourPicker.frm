VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufColourPicker 
   Caption         =   "Colour Picker"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   OleObjectBlob   =   "ufColourPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufColourPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCancelled As Boolean
Public lgColourValue As Long

Dim Buttons() As New buttonsClass


Private Sub CommandButton123_Click()
' Purpose:
' Accepts:
' Returns:
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
' Purpose:
' Accepts:
' Returns:
    Dim oControl As Control
    
    'arColours = Split( _
    '"16777215,12632319,12640511,12648447,12648384,16777152,16761024,16761087," & _
    '"14737632,8421631,8438015,8454143,8454016,16777088,16744576,16744703," & _
    '"12632256,255,33023,65535,65280,16776960,16711680,16711935," & _
    '"8421504,192,16576,49344,49152,12632064,12582912,12583104 ,4210752," & _
    '"128,16512,32896,32768,8421376,8388608,8388736,0," & _
    '"64,4210816,16448,16384,8421376,4194304,4194368," & _
    '"16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215," & _
    '"16777215,16777215,16777215,16777215,16777215,16777215,16777215,16777215", ",")

    iControlCount = 0
    For Each oControl In Me.frButtons.Controls
        If TypeName(oControl) = "CommandButton" Then
            'CLng(arColours(iControlCount))
            iControlCount = iControlCount + 1
            oControl.BackColor = ActiveWorkbook.Colors(iControlCount)
            ReDim Preserve Buttons(1 To iControlCount)
            Set Buttons(iControlCount).ButtonGroup = oControl
        End If
    Next
End Sub


