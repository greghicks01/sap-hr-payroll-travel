VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWinAPITimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' Win32 API timer class to init and kill a timer
' Win32 API contains dispatched calls based on a repetative dispatch to a callback routine (not in the class)
'    ...oooOOOooo...
' Usage
' =================
'   Set cs = New
'
'    ...oooOOOooo...

Private cpvTimerID As Long
Private cpvCallBack As String
Private cpvObject As Object

Private Sub Class_Initialize()
' Purpose:
' Accepts:
' Returns:

End Sub

Private Sub Class_Terminate()
' Purpose:
' Accepts:
' Returns:

    EndTimer
    
End Sub

Public Property Let CallBackObj(obj As Object)
' Purpose:
' Accepts:
' Returns:

    Set cpvObject = obj
    
End Property

Public Property Get CallBackObj() As Object
' Purpose:
' Accepts:
' Returns:

    Set CallBackObj = cpvObject
    
End Property

Public Property Let CallBackName(routineName As String)
' Purpose:
' Accepts:
' Returns:

     cpvCallBack = routineName
     
End Property

Public Property Get CallBackName() As String
' Purpose:
' Accepts:
' Returns:

     CallBackName = cpvCallBack
     
End Property

Public Property Let StartTimer(cpvTimerInt As Integer)
' Purpose:
' Accepts:
' Returns:

    cpvTimerID = SetTimer(0&, 0&, cpvTimerInt, AddressOf winAPITimer.TimerProc)
    
End Property

Sub EndTimer()
' Purpose:
' Accepts:
' Returns:

    On Error Resume Next
    KillTimer 0&, cpvTimerID
    
End Sub

