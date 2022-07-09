VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRRRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' Designed to build RRRS data into Analysis Results from edits, inputs and errors
'
'    ...oooOOOooo...
'
' Usage
'   set cRRRS = new clsRRRS
'
'   cRRRS.action
' =================
Private Const cpcErrorStart     As String = "Error at line "
Private Const cpcLinePos        As String = "datapool "

Private Const cpcRRRS           As String = "RRRS"
Private Const cpcDelimiter      As String = "Delimiter"
Private Const cpcEdits          As String = "Edits"
Private Const cpcErrors         As String = "Errors"
Private Const cpcFolder         As String = "Folder"
Private Const cpcInputs         As String = "Inputs"
Private Const cpcLinePrefix     As String = "Line Prefix"
Private Const cpcPadding        As String = "Padding"
Private Const cpcQuestionPrefix As String = "Question Prefix"

Private cpvDelimiter            As String
Private cpvEditFilePattern      As String
Private cpvErrorFilePattern     As String
Private cpvRootFolder           As String
Private cpvInputFilePattern     As String
Private cpvLinePrefix           As String
Private cpvPadding              As Integer
Private cpvQuestionPrefix       As String

Private cpvQLineNumber          As Integer
Private cpvResultNumber         As String

' if rootfolder is empty - open folder dialog - give option to save
Public Sub Class_Initialize()
' Purpose:
' Accepts:
' Returns:
    
    cpvDelimiter = getConfigData("Data Driven", cpcDelimiter)
    
    cpvEditFilePattern = getConfigData(cpcRRRS, cpcEdits)
    cpvErrorFilePattern = getConfigData(cpcRRRS, cpcErrors)
    cpvRootFolder = getConfigData(cpcRRRS, cpcFolder)
    cpvInputFilePattern = getConfigData(cpcRRRS, cpcInputs)
    cpvLinePrefix = getConfigData(cpcRRRS, cpcLinePrefix)
    cpvPadding = getConfigData(cpcRRRS, cpcPadding)
    cpvQuestionPrefix = getConfigData(cpcRRRS, cpcQuestionPrefix)
    
    If cpvRootFolder = "" Then
        cpvRootFolder = selectFileSystemItem(msoFileDialogFolderPicker)
        If cpvRootFolder = "-1" Then
            MsgBox "Can't perform the task", vbCritical, "No Folder Selected"
            Err.Raise 4005
        End If
    End If
    
End Sub

Public Sub Class_Terminate()
' Purpose:
' Accepts:
' Returns:
    
End Sub

Public Sub Action()
' Purpose:
' Accepts:
' Returns:

    ' set the cursor to Analysis Results C11
    Dim theSheet As String, theRange As String
    Dim Top As Integer, Right As Integer
    
    theSheet = getSheet("Analysis Start")
    theRange = getRange("Analysis Start")
    
    Worksheets(theSheet).Activate
    Range(theRange).Select
    
    'clear the existing data in C
    dimension = Application.WorksheetFunction.CountA("C:C") - 5
    
    If dimension > 0 Then
    
        Top = Range(theRange).row
        Right = Range(theRange).Column + 1
           
        Worksheets(theSheet).Range(theRange, Cells(Top + dimension, Right)).Select
        Selection.ClearContents
        
        Range(theRange).Select
    End If
    
    IterateFiles cpvEditFilePattern
    IterateFiles cpvInputFilePattern
    
    Range(theRange).Select
    IterateFiles cpvErrorFilePattern

End Sub

Private Sub IterateFiles(pattern As String)
' Purpose:
' Accepts:
' Returns:
    
    ' process edits
    Dim fso As Object, fName As Object
    Dim fileName As String
    Dim FileList As clsFileList
    Dim ReadFile As clsFileByLine
    Dim file As Variant
    
    Set FileList = New clsFileList
    Set ReadFile = New clsFileByLine
    Set fso = CreateObject("Scripting.filesystemobject")
    
    FileList.AddFiles cpvRootFolder, pattern
        
    For Each file In FileList.FileList
    
        ReadFile.file = file
        
        fileName = fso.GetFileName(file)
                
        'always skip first line
        line = ReadFile.NextLine
        cpvQLineNumber = 0
        
        While Not ReadFile.EndOfData
        
            line = ReadFile.NextLine
            cpvQLineNumber = cpvQLineNumber + 1
            
            Select Case pattern
                Case Is = cpvEditFilePattern
                    ProcessEdits line
                    
                Case Is = cpvInputFilePattern
                    ProcessInputs fileName, line
                    
                Case Is = cpvErrorFilePattern
                    ProcessErrors line
            End Select
            
            DoEvents
            
        Wend
    Next
    
    Set FileList = Nothing
    Set ReadFile = Nothing
    
End Sub

Private Sub ProcessEdits(ByVal Data As String)
' Purpose:
' Accepts:
' Returns:

    Dim TestCaseID As String
    Dim csv As clsCSV
    
    If InStr(1, Data, "quit") <> 0 Then Exit Sub
    
    Set csv = New clsCSV
    csv.csvString = Data
    
    ' Question + Line + message
    ' 1,3,15,1,verify,"Primary payment is not cancelled in Q1"
    
    TestCaseID = cpvQuestionPrefix & PadNumber(csv.columnN(1), cpvPadding) & cpvDelimiter & "Edits" & cpvDelimiter
    TestCaseID = TestCaseID & cpvLinePrefix & PadNumber(CStr(cpvQLineNumber), cpvPadding) & cpvDelimiter
    TestCaseID = TestCaseID & csv.columnN(csv.csvColumn("verify") + 1)
    
    ' write to results
    writeData TestCaseID
    'write line and pass to results

End Sub

Private Sub ProcessInputs(ByVal fileName As String, ByVal Data As String)
' Purpose:
' Accepts:
' Returns:

    Dim TestCaseID As String
    Dim csv As clsCSV
    
    If InStr(1, Data, "quit") <> 0 Then Exit Sub
    
    Set csv = New clsCSV
    csv.csvString = Data
    
    ' Question + Line - pull the value after verify
    TestCaseID = cpvQuestionPrefix & PadNumber(Mid(fileName, 1, InStr(1, fileName, "_") - 1), cpvPadding) & cpvDelimiter & "Inputs" & cpvDelimiter
    TestCaseID = TestCaseID & cpvLinePrefix & PadNumber(CStr(cpvQLineNumber), cpvPadding) & cpvDelimiter
    
    TestCaseID = TestCaseID & Data
    
    'write line and pass to results
    writeData TestCaseID
    
    Set csv = Nothing
    
End Sub

Private Sub writeData(ByVal TestCaseID As String)
' Purpose:
' Accepts:
' Returns:

    ' write to col C
    ActiveCell.value = TestCaseID
    
    'advance to col d
    ActiveCell.Offset(0, 1).Select
    ActiveCell.value = "p"
    
    'advance down and left to next row in C
    ActiveCell.Offset(1, -1).Select

End Sub

Private Sub ProcessErrors(ByVal Data As String)
' Purpose:
' Accepts:
' Returns:

' Error at line 10 in datapool 33_edits.csv:
'               ^^ Line number ^^ Q number
' Q33 is answered 'Other' but Q34 not answered NOT FOUND IN ERROR WINDOW
' Error Message
    Dim splitPoint As Integer, FrontString As String, ErrorString As String, temp As String, TestCaseID As String

    If InStr(1, Data, cpcErrorStart) <> 1 Then Exit Sub
    
    splitPoint = InStr(1, Data, ":")
    FrontString = Mid(Data, 1, splitPoint - 1)
    ErrorString = Mid(Data, splitPoint + 1)
    
    splitPoint = InStr(1, FrontString, cpcLinePos) + Len(cpcLinePos)
    temp = Mid(FrontString, splitPoint)
    TestCaseID = cpvQuestionPrefix & PadNumber(Left(temp, InStr(1, temp, "_") - 1), cpvPadding)
    temp = Mid(FrontString, InStr(1, FrontString, "_") + 1)
    temp = Left(temp, InStr(1, temp, ".") - 1)
    TestCaseID = TestCaseID & cpvDelimiter & temp
    
    splitPoint = InStr(1, FrontString, cpcErrorStart) + Len(cpcErrorStart)
    temp = Mid(FrontString, splitPoint)
    TestCaseID = TestCaseID & cpvDelimiter & cpvLinePrefix & PadNumber(Left(temp, InStr(1, temp, " ") - 1), cpvPadding)
        
    ' write f to relevant q
    UpdateFails TestCaseID
    
End Sub

Private Sub UpdateFails(ByVal Search As String)
' Purpose:
' Accepts:
' Returns:

    ' search for wildcarded data
    On Error GoTo searchNotFound
    
        Cells.Find(What:=Search & "*", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
        ActiveCell.Offset(0, 1).Select
        ActiveCell.value = "f"
        
    Exit Sub
    
searchNotFound:

End Sub


