VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReportGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' Designed to manage report generation
'  Collects saved pages for report
'
'    ...oooOOOooo...
'
' Usage
' =================
'
Option Explicit

' names of saved data pages for a report
Private cpvReportPages  As Collection
Private cpvSavedList    As Collection

Public Sub Class_Initialize()
' Purpose:
' Accepts:
' Returns:
    Set cpvSavedList = New Collection
    Set cpvReportPages = New Collection
    
End Sub

Public Sub Class_Terminate()
' Purpose:
' Accepts:
' Returns:
    Set cpvSavedList = Nothing
    Set cpvReportPages = Nothing
    
End Sub

Public Property Let addResults(ByVal rName As String)
' Purpose:
' Accepts:
' Returns:
    cpvSavedList.Add rName
    cpvReportPages.Add "Analysis Results " & rName
    
End Property
