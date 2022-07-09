VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAGSgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class AGS Num
' P: Sequentially Generates a valid AgS number for use on the master sheet
' A: nil - values are stored in spare sheet cells A1 and A2
' R:
Private startV As Long
Private stopV As Long
Private AGSNum As String
Private fso As Object
Private Const fileName = "agsnum.txt"


Private Sub Class_Initialize()
' P:
' A:
' R:

' read from a file to protect from accidental
    startV = Sheets("BSB").Range("L1")
    stopV = Sheets("BSB").Range("L2")

End Sub

Private Sub Class_Terminate()
' P:
' A:
' R:
' write to file
End Sub


Function agsChecksum(ByVal AGSNum As Long) As Integer
' P:
' A:
' R:

    agsStr = StrReverse(CStr(AGSNum))
    agsChecksum = 0
    
    For x = 1 To 7
        temp = Mid(agsStr, x, 1)
        agsChecksum = agsChecksum + CInt(temp) * 2 ^ x
    Next
    
    agsChecksum = agsChecksum Mod 11
    
    If agsChecksum <> 0 Then agsChecksum = 11 - agsChecksum

End Function

Public Property Get NextAGS() As String
' P:
' A:
' R:
Dim temp As Integer
   
    Do
        startV = startV + 1
        temp = agsChecksum(startV)
    Loop While temp = 10
    
    NextAGS = CStr(startV) & temp
    Sheets("BSB").Range("L1") = startV
      
End Property
