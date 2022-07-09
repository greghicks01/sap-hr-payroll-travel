Attribute VB_Name = "PublicTypes"
Option Explicit

Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const MAX_PATH As Integer = 260

Public Const SPath = "C:\Automation\SAPQTP\QC Project\Test Resources\Data Tables"
'Public Const sExpPath = "C:\2. Test Preparation\01 Automation\QC Project\GUI\Data"
'public const sMyName="SAPConsolRecords.xlsm"

Public oNameCounter As clsNameGen

Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * 14
End Type

Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" _
  Alias "GetUserNameA" ( _
  ByVal lpBuffer As String, _
  ByRef nSize As Long) As Long
  
Public Declare PtrSafe Function GetComputerName Lib "kernel32.dll" _
  Alias "GetComputerNameA" ( _
  ByVal lbbuffer As String, _
  nSize As Long) As Long
    
Public Declare PtrSafe Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    
Public Declare PtrSafe Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    
Public Declare PtrSafe Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Public Declare PtrSafe Function OpenClipboard Lib "User32.dll" (ByVal hWndNewOwner As Long) As Long
  
Public Declare PtrSafe Function EmptyClipboard Lib "User32.dll" () As Long

Public Declare PtrSafe Function CloseClipboard Lib "User32.dll" () As Long

Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long
    
Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal nIDEvent As Long) As Long
    
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Sub setXLFocus()
' P:
' A:
' R:

    Dim xlHdl As Long

    xlHdl = FindWindow("XLMAIN", Application.Caption)
    SetForegroundWindow xlHdl

End Sub

Public Function TrimNull(sFileName As String) As String
' P:
' A:
' R:
    Dim i As Long
    ' Search for the first null character
    i = InStr(1, sFileName, vbNullChar)
    If i = 0 Then
        TrimNull = sFileName
    Else
        ' Return the file name
        TrimNull = Left$(sFileName, i - 1)
    End If
End Function

Public Function PadNumber(Number As String, pad As Integer) As String
' P:
' A:
' R:
    Dim i As Integer
    i = pad - Len(Trim(Number))
    PadNumber = String(IIf(i >= 0, i, 0), "0") + Trim(Number)
End Function

Public Function getConfigData(ConfigItem As String) As String
' Purpose
' Accepts
' Returns
    Dim theSheet As String, _
        theRange As String
        
    theSheet = "Configuration"
    theRange = "A2:B13"
    getConfigData = Application.WorksheetFunction.VLookup(ConfigItem, Worksheets(theSheet).Range(theRange), 2)
    
End Function

Public Sub setConfigData(ConfigArea As String, ConfigItem As String, Data As String)
' Purpose Sets the data in a range and column,,,,
' Accepts
' Returns
    Dim ARange As Range, _
        theSheet As String, _
        theRange As String, _
        LeftColNum As Integer, _
        RightColNum  As Integer, _
        TopRowNum As Integer, _
        BotRowNum  As Integer, _
        RelRow As Integer

    'theSheet = getSheet(ConfigArea)
    'theRange = getRange(ConfigArea)
    
    Set ARange = Range(theRange)
    
    LeftColNum = ARange.column
    RightColNum = ARange.Columns(ARange.Columns.Count).column
    TopRowNum = ARange.Row
    BotRowNum = ARange.Rows(ARange.Rows.Count).Row
    
    theRange = Range(Cells(TopRowNum, LeftColNum), Cells(BotRowNum, LeftColNum)).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    
    'Row/Column
    'RelRow = Application.WorksheetFunction.Match(ConfigItem, Worksheets(theSheet).Range(theRange), 0) - 1
    
    ' write the data out to the Cell RelRow from TopRowNum,RightColNum
    'Worksheets(theSheet).Cells(TopRowNum + RelRow, RightColNum).Value = Data
    
End Sub

Public Sub ClearClipboard()
' Purpose:
' Accepts:
' Returns:

  Dim Ret
  
    Ret = OpenClipboard(0&)
      If Ret <> 0 Then Ret = EmptyClipboard
    CloseClipboard
    
End Sub
