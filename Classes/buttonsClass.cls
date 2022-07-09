VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "buttonsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents ButtonGroup As CommandButton
Attribute ButtonGroup.VB_VarHelpID = -1

Private Sub ButtonGroup_MouseDown(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

'   Assign the value and close the dialog
    ufColourPicker.lgColourValue = ButtonGroup.BackColor
    ufColourPicker.Hide
End Sub

Private Sub ButtonGroup_MouseMove _
    (ByVal Button As Integer, ByVal Shift As Integer, _
     ByVal x As Single, ByVal Y As Single)
'   Update the sample in the dialog
    Dim HexDigits As String
    Dim BluePart As Integer, GreenPart As Integer, RedPart As Integer
    ufColourPicker.SAMPLELabel.BackColor = ButtonGroup.BackColor
    HexDigits = Hex(ButtonGroup.BackColor)
    Do Until Len(HexDigits) = 6
        HexDigits = "0" & HexDigits 'pad with zeros
    Loop
    BluePart = Val("&h" & Left(HexDigits, 2))
    GreenPart = Val("&h" & Mid(HexDigits, 3, 2))
    RedPart = Val("&h" & Right(HexDigits, 2))
    ufColourPicker.RGBLabel.Caption = RedPart & " " & GreenPart & " " & BluePart
    
End Sub

