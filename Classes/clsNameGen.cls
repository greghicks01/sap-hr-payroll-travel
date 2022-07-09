VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNameGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public oDic As Object

Private Sub Class_Initialize()
    Set oDic = CreateObject("scripting.dictionary")
End Sub

Private Sub Class_Terminate()
    Set oDic = Nothing
End Sub
