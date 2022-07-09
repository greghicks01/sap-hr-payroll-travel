VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDataDriven"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' Designed to process csv data files and generate a line in the results based on the Data driven configuration
'
'    ...oooOOOooo...
'
' Usage
' =================
'

' Purpose:
' Accepts:
' Returns:
Option Explicit

Private Const cpcIndex      As String = "Analysis Start"
Private Const cpcData       As String = "Data Driven"
Private Const cpcConfigList As String = "columns,delimiter,File,Quit Tag"

Private cpvFileRead         As clsFileByLine
Private cpvColumnData       As clsCSV

Private cpvPadding          As Integer
Private cpvLineCount        As Integer

Private cpvTargetSht        As String
Private cpvtargetRng        As String
Private cpvAddDelim         As String
Private cpvFileName         As String
Private cpvQuitTag          As String

Public Sub Class_Initialize()
' Purpose:
' Accepts:
' Returns:
    Dim cDataDriven As clsCSV
    
    Set cDataDriven = New clsCSV
    Set cpvColumnData = New clsCSV
    
    cpvPadding = 3
    
    cDataDriven.csvString = cpcConfigList
    
    ' get the config elements
    cpvColumnData.csvString = getConfigData(cpcData, cDataDriven.columnN(1))
    cpvAddDelim = getConfigData(cpcData, cDataDriven.columnN(2))
    cpvFileName = getConfigData(cpcData, cDataDriven.columnN(3))
    cpvQuitTag = UCase(getConfigData(cpcData, cDataDriven.columnN(4)))
    
    cpvTargetSht = getSheet(cpcIndex)
    cpvtargetRng = getRange(cpcIndex)
    
    If cpvFileName = "" Then
        cpvFileName = getFile
    End If
    
    Set cDataDriven = Nothing
    
    If cpvFileName = "-1" Then Err.Raise 4004
    
End Sub

Public Sub Class_Terminate()
' Purpose:
' Accepts:
' Returns:
    Set cpvColumnData = Nothing
    
End Sub

Private Function getFile() As String
' Purpose:
' Accepts:
' Returns: Filename


    getFile = selectFileSystemItem(msoFileDialogOpen, "*.csv")
     
End Function

Public Sub Action()
' Purpose:
' Accepts:
' Returns:
    Dim temp As String

    Worksheets(cpvTargetSht).Activate
    Worksheets(cpvTargetSht).Range(cpvtargetRng).Select

    Set cpvFileRead = New clsFileByLine
    
    cpvFileRead.file = cpvFileName
    
    ' skip first line
    temp = cpvFileRead.NextLine
    
    While Not cpvFileRead.EndOfData
        processDMISCSV cpvFileRead.NextLine
    Wend

 End Sub
 
Private Sub processDMISCSV(line As String)
' Purpose: tuans DMIS stuff
' Accepts:
' Returns:

    Dim cDataSource As clsCSV, _
        ResultStr   As String, _
        element     As Variant
     
    ' quit and skip lines here
    If InStr(1, UCase(line), cpvQuitTag) <> 0 Then Exit Sub
    
    Set cDataSource = New clsCSV 'ConfigListCol

    cDataSource.csvString = line
    cpvLineCount = cpvLineCount + 1
    ResultStr = ""
    
    For Each element In cpvColumnData.csvCollection
        Select Case UCase(element)
            Case "[ITER]"
                ResultStr = PadNumber(ResultStr & Trim(Str(cpvLineCount)), cpvPadding)
            Case Else
                ResultStr = ResultStr & cDataSource.columnA(element)
        End Select
        ' add custom delimeter
        ResultStr = ResultStr & cpvAddDelim
        
    Next
        
    ResultStr = Mid(ResultStr, 1, Len(ResultStr) - 1)
    ActiveCell.value = ResultStr ' write data to cell
    ActiveCell.Offset(1, 0).Select
    DoEvents
    
    Set cDataSource = Nothing
    
End Sub



