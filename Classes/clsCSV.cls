VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' Designed to manage comma separated data
' hand a CSV line in by setting  csvString, get it back by reading csvString
' get an element via a numerical reference or Alpha Column reference (columnN or columnA)
' get or set the collection item by using csvColl
' find the column number by using csvCol
'    ...oooOOOooo...
' Usage
' =================
'    Set cs = New clsCSV
'
'    cs.csvString = "absolute,best,charmed"
'    Debug.Print cs.columnN(1)
'    Debug.Print cs.columnA("B")
'
'    Debug.Print cs.columnA("AB")
'
'    Debug.Print cs.csvString
'    ...oooOOOooo...

Option Explicit

Private cpvDelimiter    As String
Private cpvCSVCol       As Collection

Public Sub Class_Initialize()
' Purpose:
' Accepts:
' Returns:

    cpvDelimiter = ","
    Set cpvCSVCol = New Collection
    
End Sub

Public Sub Class_Terminate()
' Purpose:
' Accepts:
' Returns:

    Set cpvCSVCol = Nothing
    
End Sub

Public Property Let csvAddCollection(c As Collection)
' Purpose: adds new collection to the private collection
' Accepts:
' Returns:
    Dim a As Variant

    For Each a In c
    
        cpvCSVCol.Add a
        
    Next a

End Property

Public Property Let csvCollection(c As Collection)
' Purpose: clears and sets the new collection
' Accepts:
' Returns:

    If cpvCSVCol.Count > 0 Then
        
        Set cpvCSVCol = Nothing
        Set cpvCSVCol = New Collection
        
    End If
    
    Set cpvCSVCol = c

End Property

Public Property Get csvCollection() As Collection
' Purpose:
' Accepts:
' Returns:

        Set csvCollection = cpvCSVCol

End Property

Public Property Let csvString(ByVal rhs As String)
' Purpose:
' Accepts:
' Returns:
    Dim x As Integer
    Dim c As Variant
        
    If cpvCSVCol Is Nothing Then Set cpvCSVCol = New Collection
    
    If cpvCSVCol.Count > 0 Then
        
        Set cpvCSVCol = Nothing
        Set cpvCSVCol = New Collection
        
    End If
    
    For x = 1 To countDelim(rhs) + 1
        If InStr(1, rhs, cpvDelimiter) <> 0 Then
            cpvCSVCol.Add Mid(rhs, 1, InStr(1, rhs, cpvDelimiter) - 1)
            rhs = Mid(rhs, InStr(1, rhs, cpvDelimiter) + 1)
            
        Else
            cpvCSVCol.Add rhs
            
        End If
    Next
    
End Property

Public Property Get csvString() As String
' Purpose:
' Accepts:
' Returns:
    Dim x As Integer
    
    csvString = cpvCSVCol.Item(1)
    
    For x = 2 To cpvCSVCol.Count
    
        csvString = csvString + cpvDelimiter + cpvCSVCol.Item(x)
            
    Next
        
End Property

Public Property Get csvColumn(value As String) As Integer
' Purpose:
' Accepts:
' Returns:
    Dim x As Integer
    
    For x = 1 To cpvCSVCol.Count
    
        If cpvCSVCol.Item(x) = value Then
        
            csvColumn = x
            Exit Property
            
        End If
        
    Next
    
End Property

Public Property Get columnN(ByVal n As Integer) As String
' Purpose:
' Accepts:
' Returns:

    columnN = cpvCSVCol.Item(n)
        
End Property

Public Property Get columnA(ByVal CellColumn As String) As String
' Purpose: Convert cell column references to CSV element
' Accepts: Column Letters, A-Z,AA-ZZ etc
' Returns:
    Dim x As Integer
    Dim sum As Integer
    
    ' reverse the string for the calculation
    CellColumn = StrReverse(CellColumn)
    
    ' loop calculates each position
    sum = 0
    
    For x = 1 To Len(CellColumn)
    
        sum = sum + ((Asc(Mid(CellColumn, x, 1)) - 64) * (26 ^ (x - 1)))
        
    Next
    
    columnA = cpvCSVCol.Item(sum)
    
End Property

Private Function countDelim(ByVal csv As String) As Integer
' Purpose: takes a CSV data string and counts the delimiters
' Accepts: csv
' Returns: delim count

    countDelim = 0
    
    Do Until InStr(1, csv, cpvDelimiter) = 0
    
        countDelim = countDelim + 1
        csv = Mid(csv, InStr(1, csv, cpvDelimiter) + 1)
        
    Loop
    
End Function
