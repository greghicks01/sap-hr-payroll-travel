VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufColumnRestructure 
   Caption         =   "Column Restructure"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   OleObjectBlob   =   "ufColumnRestructure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufColumnRestructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bCancelled As Boolean
Private oDicList As Object
Private srcWb As Workbook, srcWs As Worksheet
Private Const srcSheet = "Default Data"

Private Sub cbColour_Click()
' Purpose:
' Accepts:
' Returns:

    If Me.lbHeaderNames.ListIndex > 0 Then
        'find colour and show selector
        ColorDialog
    End If

End Sub

Private Sub cbRename_Click()
' Purpose: Enables columns to be renamed during restructure
' Accepts:
' Returns:

    Dim sNewname As String

    If Me.lbHeaderNames.ListIndex = -1 Then
        Exit Sub
    End If

    If MsgBox("Agreement between Automation and PST obtained?", vbYesNo, "Change Column Name") = vbNo Then
        Exit Sub
    End If
    
    tmp = lbHeaderNames.ListIndex
    
    ' display the rename dlg
    sNewname = InputBox("Update the name", "Change Header", lbHeaderNames.List(tmp))
    
    ' confirm name is not duplicate
    If locateHeader(ActiveSheet, sNewname) <> 0 Then
        MsgBox "Cannot complete the update", vbOKOnly, "Duplicate Name"
        Exit Sub
    End If
    ' change the header row cell
    lbHeaderNames.List(tmp) = sNewname
    ActiveSheet.Cells(1, tmp + 1) = sNewname

End Sub

Private Sub cbUp_Click()
' Purpose: Moves item up in the list and left on the sheet
' Accepts:
' Returns:

    If Me.lbHeaderNames.ListIndex > 0 Then
        tmp = lbHeaderNames.ListIndex
        ActiveSheet.Columns(tmp + 1).Select
        Selection.Cut
        
        ActiveSheet.Columns(tmp).Select
        Selection.Insert Shift:=xlToRight
        
        Call buildList
        
        lbHeaderNames.ListIndex = tmp - 1
        
    End If
    
End Sub

Private Sub cbDown_Click()
' Purpose: Moves item down in the list and right on the sheet
' Accepts:
' Returns:

    If Me.lbHeaderNames.ListIndex >= 0 And Me.lbHeaderNames.ListIndex < Me.lbHeaderNames.ListCount - 1 Then
        
        tmp = lbHeaderNames.ListIndex
        ActiveSheet.Columns(tmp + 1).Select
        Selection.Cut
        
        ActiveSheet.Columns(tmp + 3).Select
        Selection.Insert Shift:=xlToRight
        
        Call buildList
        
        lbHeaderNames.ListIndex = tmp + 1
        
    End If
    
End Sub

Private Sub cbDelete_Click()
' Purpose: Removes column from temp sheet
' Accepts:
' Returns:

' update the persistant object

    If Me.lbHeaderNames.ListIndex >= 0 Then
        ActiveSheet.Columns(Me.lbHeaderNames.ListIndex + 1).Select
        If MsgBox("Are you Sure?", vbYesNo, "Delete Selected Column") = vbYes Then
            Selection.Delete Shift:=xlToLeft
            Call buildList
        End If
    End If

End Sub

Private Sub cbInsert_Click()
' Purpose: Inserts column before the selected one on the temp sheet
' Accepts:
' Returns:

    Dim colName As String

    If Me.lbHeaderNames.ListIndex >= 0 Then
        ActiveSheet.Columns(Me.lbHeaderNames.ListIndex + 1).Select
        If MsgBox("Are you Sure?", vbYesNo, "Insert before selected column") = vbYes Then
            'Name of column
            colName = InputBox("Please enter a column name (must be unique)")
            If colName <> "" Or InStr(1, colName, " ") <> 0 Then 'no blanks
                If locateHeader(ActiveSheet, colName) = 0 Then
                    Selection.Insert Shift:=xlToRight
                                        
                    ufColourPicker.Show
    
                    If ufColourPicker.bCancelled = True Then
                        Exit Sub
                    End If
                    
                    'add name to inserted column
                    ActiveSheet.Cells(1, Me.lbHeaderNames.ListIndex + 1) = colName
                    ActiveSheet.Cells(1, Me.lbHeaderNames.ListIndex + 1).Interior.Color = ufColourPicker.lgColourValue
                    
                    tmp = Me.lbHeaderNames.ListIndex
                    ' color picker
                    Call buildList
                    Me.lbHeaderNames.ListIndex = tmp
                Else
                    If colName = "" Then
                        sHead = "a blank"
                    Else
                        sHead = "spaces in"
                    End If
                    MsgBox "Cannot accept " & sHead & " name", vbOKCancel, "Name Error"
                End If
            Else
                MsgBox "Cancelled"
            End If
        End If
    End If

End Sub

Private Sub cbOK_Click()
' Purpose: Complete
' Accepts:
' Returns:

    ' drag it all back in import from temp to master
    If MsgBox("", vbYesNo, "Proceed") = vbYes Then
        Me.Hide
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        
        update_default_data
        import_column_change
        Unload Me
        
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End If
    
End Sub

Private Sub update_default_data()
' Purpose: Updates the default data page
' Accepts:
' Returns:
' oDicList, lbHeaderNames

' Change to use the new object in the main sheet
' will reduce potential errors

    For el = 0 To lbHeaderNames.ListCount - 1
        srcWs.Cells(el + 2, "A") = el + 1
        srcWs.Cells(el + 2, "B") = lbHeaderNames.List(el, 0)
        If oDicList.exists(lbHeaderNames.List(el, 0)) Then
            arTmp = Split(oDicList.Item(lbHeaderNames.List(el, 0)), "|")
            iTmpCol = 3
            For Each elem In arTmp
                If InStr(1, elem, "0") = 1 Then elem = "'" + elem
                srcWs.Cells(el + 2, iTmpCol) = elem
                iTmpCol = iTmpCol + 1
            Next
        Else
            For x = 3 To 5
                srcWs.Cells(el + 2, iTmpCol) = ""
            Next
        End If
    Next
End Sub

Private Sub cbCancel_Click()
' Purpose: Moves item up and down in the list and on the sheet
' Accepts:
' Returns:

    bCancelled = True
    Unload Me
    On Error Resume Next
    Workbooks("temp.xlsx").Close False
    On Error GoTo 0
End Sub


Private Sub UserForm_Initialize()
' Purpose: get the header names off the temp file
' Accepts:
' Returns:

    
    
    Set srcWb = Workbooks(1)
    Set srcWs = srcWb.Sheets(srcSheet)
    
    Call buildList
    
    Set oDicList = CreateObject("scripting.dictionary")
    
    For nCol = 0 To lbHeaderNames.ListCount - 1
        oDicList.Item(CStr(lbHeaderNames.List(nCol))) = srcWs.Cells(nCol + 2, "C") & "|" & srcWs.Cells(nCol + 2, "D") & "|" & srcWs.Cells(nCol + 2, "E")
    Next
    
    bCancelled = False
    
End Sub

Private Sub buildList()
' Purpose:
' Accepts:
' Returns:
    
    Me.lbHeaderNames.Clear

    With ActiveSheet
        nCol = 1
        While ActiveSheet.Cells(1, nCol) <> ""
            Me.lbHeaderNames.AddItem CStr(.Cells(1, nCol))
            nCol = nCol + 1
        Wend
    End With

End Sub

