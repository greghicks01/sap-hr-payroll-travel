VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' The timer class is to enable routine saves and shutdowns to occur.
'    ...oooOOOooo...
' Usage
' =================
'   Set cs = New clsTimer - sets default time of 30 minutes
'
'   'Use these for other intervals
'   cs.SecondsInterval = value
'   cs.MinutesInterval = value (default is 30 minutes)
'   cs.HoursInterval = value
'
'   cs.StartTimer = "myProc" ' myproc is the string name of the called procedure
'
'   set cs = nothing ' auto terminates timer
'
'   'Inside your procedure include this
'   cs.Restart
'
'    ...oooOOOooo...

Option Explicit

Private cpvRunIntervalSec As Integer
Private cpvRunIntervalMin As Integer
Private cpvRunIntervalHrs As Integer

Private cpvRunWhen As Double
Private cpvRunWhat As String

Private Sub Class_Initialize()
' Purpose:
' Accepts:
' Returns:

End Sub

Private Sub Class_Terminate()
' Purpose:
' Accepts:
' Returns:

    StopTimer

End Sub

Property Let RunWhat(rhs As String)
' Purpose:
' Accepts:
' Returns:

    cpvRunWhat = rhs

End Property

Property Let SecondsInterval(rhs As Integer)
' Purpose:
' Accepts:
' Returns:

    cpvRunIntervalSec = rhs

End Property

Property Let MinutesInterval(rhs As Integer)
' Purpose:
' Accepts:
' Returns:

    cpvRunIntervalMin = rhs

End Property

Property Let HoursInterval(rhs As Integer)
' Purpose:
' Accepts:
' Returns:

    cpvRunIntervalHrs = rhs

End Property

Public Property Let StartTimer(ProcName As String)
' Purpose:
' Accepts:
' Returns:

    cpvRunWhat = ProcName
    
    Restart

End Property

Public Sub Restart()
' Purpose:
' Accepts:
' Returns:
    
    cpvRunWhen = Now + TimeSerial(cpvRunIntervalHrs, cpvRunIntervalMin, cpvRunIntervalSec)

    Application.OnTime EarliestTime:=cpvRunWhen, procedure:=cpvRunWhat, Schedule:=True

End Sub

Public Sub StopTimer()
' Purpose:
' Accepts:
' Returns:

    On Error Resume Next
    Application.OnTime EarliestTime:=cpvRunWhen, procedure:=cpvRunWhat, Schedule:=False

End Sub

