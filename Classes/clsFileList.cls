VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' Designed to recursively build a list of folders and files (full path)
' Start by assigning property RootFolder, then FileSearch with wildcarded parameters
' Use NextFile or NextFolder in a do/while loop
'    ...oooOOOooo...
'
' Usage
' =================
'
'   Dim FL as clsFileList
'   set FL = new clsFileList
'   FL.rootFolder = "C:\Data\Workspace"
'   FL.FileSearch = "*.csv"
'
'   Set c = FileList.FileList
'
'    For Each it In c
'       ' put your operations here using it to read each line of the file eg
'       Debug.Print it
'   wend
'

Option Explicit

Private cpvFso          As Object
Private cpvFldObj       As Object
Private cpvFiles        As Collection
Private cpvFolders      As Collection

Private cpvRootFolder   As String
Private cpvWildCard     As String

Public Sub Class_Initialize()
' Purpose: Class event on New or create
' Accepts:
' Returns:

    Set cpvFso = CreateObject("Scripting.FileSystemObject")
    Set cpvFiles = New Collection
    Set cpvFolders = New Collection
    cpvRootFolder = "C:\"
    
End Sub

Public Sub Class_Terminate()
' Purpose: Class event on Nothing or loss of scope
' Accepts:
' Returns:
    Set cpvFso = Nothing
    Set cpvFiles = Nothing
    Set cpvFolders = Nothing
    Set cpvFldObj = Nothing
    
    cpvRootFolder = ""
    
End Sub

Public Sub AddFiles(ByVal path As String, Optional ByVal wildcard As String)
' Purpose: accepts a wildcard and a root path to find files in folder and subFolders
' Accepts:
' Returns:
    RootFolder = path
    FileSearch = wildcard
    
End Sub

Public Property Get FileList() As Collection
' Purpose:
' Accepts:
' Returns:
    Dim obItem As Variant
    Set FileList = New Collection
    
    For Each obItem In cpvFiles
        FileList.Add obItem
    Next
    
End Property

Public Property Get FolderList(ByRef c As Collection)
' Purpose:
' Accepts:
' Returns:
    Dim obItem As Object
    Set FolderList = New Collection
    
    For Each obItem In cpvFolders
        c.Add obItem
    Next
    
End Property

Public Property Get RootFolder() As String
' Purpose:
' Accepts:
' Returns:
    RootFolder = cpvRootFolder
        
End Property

Public Property Let RootFolder(ByVal nVal As String)
' Purpose:
' Accepts:
' Returns:

    cpvRootFolder = nVal
        
End Property

Private Property Let FileSearch(ByVal vNewVal As String)
' Purpose:
' Accepts:
' Returns:
    Dim cpvSearchData As WIN32_FIND_DATA, _
        fso As Object, _
        fsfolders As Object, _
        fsFolder As Object
    
    Set cpvFiles = Nothing
    Set cpvFiles = New Collection
  
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fsfolders = fso.GetFolder(cpvRootFolder)
    Set fsfolders = fsfolders.Subfolders
    
    cpvSearchData.cFileName = vNewVal
    
    findAllFiles cpvRootFolder, vNewVal, cpvSearchData
    cpvFolders.Add cpvRootFolder
    
    For Each fsFolder In fsfolders
    
        findAllFolders fsFolder, vNewVal, cpvSearchData
        cpvFolders.Add fsFolder.path
        
    Next

End Property

Private Sub findAllFolders(Folder As Object, vNewVal As String, sdType As WIN32_FIND_DATA)
' Purpose:
' Accepts:
' Returns:

    Dim f As Object
        
    findAllFiles Folder.path, vNewVal, sdType
    cpvFolders.Add Folder.path
    
    For Each f In Folder.Subfolders
        DoEvents
        findAllFolders f, vNewVal, sdType
        cpvFolders.Add Folder.path
        
    Next

End Sub

Private Sub findAllFiles(path As String, wildcard As String, searchData As WIN32_FIND_DATA)
' Purpose: win32 implementation of wildcarded search parameters
' Accepts:
' Returns:

    Dim iSearchHandle As Long, _
        fso As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
            
    iSearchHandle = FindFirstFile(fso.buildPath(path, wildcard), searchData)
    
    If iSearchHandle <> INVALID_HANDLE_VALUE Then
    
        Do
            cpvFiles.Add fso.buildPath(path, TrimNull(searchData.cFileName))
        Loop Until FindNextFile(iSearchHandle, searchData) = 0
        
    End If
    
    FindClose (iSearchHandle)
    
    Set fso = Nothing
    
End Sub


