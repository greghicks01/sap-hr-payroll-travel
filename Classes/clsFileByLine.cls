VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileByLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' Designed to read a file line by line
' Start by setting property file with the file name
' Use in a do/while loop to iterate each line
'
'    ...oooOOOooo...
'
' Usage
' =================
'
'   Dim FL as clsFileByLine
'   set FL = new clsFileByLine
'   FL.File = <fullPathandFileName>
'
'   while Not FL.EndOfData
'       ' put your operations here using FL.NextLine to read each line of the file eg
'       Debug.Print FL.NextLine
'   wend
'

Option Explicit

Private cflFso  As Object
Private cflTs   As Object

Public Sub Class_Initialize()
' Purpose:
' Accepts:
' Returns:

    Set cflFso = CreateObject("Scripting.FileSystemObject")
        
End Sub

Public Sub Class_Terminate()
' Purpose:
' Accepts:
' Returns:

    Set cflFso = Nothing
    Set cflTs = Nothing
        
End Sub

Public Property Let file(ByVal fileName As String)
' Purpose: opens file for reading
' Accepts: fullPath and filename
' Returns:
    'If Not cflTs Is Nothing Then cflTs.Close
    
    Set cflTs = cflFso.OpenTextFile(fileName, 1)

End Property

Public Property Get EndOfData() As Boolean
' Purpose: used to detect end of file in stream
' Accepts:
' Returns:

    EndOfData = cflTs.AtEndOfStream

End Property

Public Property Get NextLine() As String
' Purpose: returns next line without errors from overflow EOF etc
' Accepts:
' Returns:
    
    NextLine = cflTs.ReadLine
            
End Property


