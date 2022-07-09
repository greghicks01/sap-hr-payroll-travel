Attribute VB_Name = "UserChangeLog"
' insert a row and add the data
Private Const logSheet As String = "Log"
Private Const rowEntry = 3

Sub logUsersChanges(Target As Range, page As String)
' Purpose:
' Accepts:
' Returns:
    
    Dim oldVal As Collection
    Dim newLoc As String
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    newLoc = ActiveCell.Address
    
    Call ClearClipboard
    
    Set oldVal = New Collection
    
    On Error Resume Next
     
    With Application
    
        .EnableEvents = False
        .Undo
    
        For Each c In Target
            oldVal.Add c.Value
        Next
    
        .Undo
        .EnableEvents = True
        
    End With
    
    For Each c In Target
        
        ' subv here
        Call logHeader
        
        'F Sheet Name
        Worksheets(logSheet).Range("F" & rowEntry).Value = page
        
        ' G = Script Name from Col & 3
        Select Case c.column
            Case Is = 4, 5
                Worksheets(logSheet).Range("G" & rowEntry).Value = Worksheets(Target.Worksheet.name).Cells(3, 3).Value
            Case Else
                Worksheets(logSheet).Range("G" & rowEntry).Value = Worksheets(Target.Worksheet.name).Cells(3, c.column).Value
        End Select
        
        ' h = Dataset Name from D & Row
        Worksheets(logSheet).Range("H" & rowEntry).Value = Worksheets(Target.Worksheet.name).Cells(c.Row, 4)
        'Target V
        Worksheets(logSheet).Range("I" & rowEntry).Value = Target.Address(RowAbsolute:=False, ColumnAbsolute:=False)
        
        'H Data Before
        Worksheets(logSheet).Range("J" & rowEntry).Value = oldVal.Item(1)
        oldVal.Remove 1
        
        'J Data After
        Worksheets(logSheet).Range("K" & rowEntry).Value = c.Value
        
        Worksheets(logSheet).Range("L" & rowEntry).Value = InputBox("Please comment on change")
    
    Next
    
    Set oldVal = Nothing
    
    Worksheets(page).Range(newLoc).Activate
    
    On Error GoTo 0
      
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Sub writeProgress()
' Purpose:
' Accepts:
' Returns:

    PathName = "C:\SAP consolidation\Test management\2. Test Preparation\01 Automation\Exported Payrolls\Progress.xlsx"
    PageNames = "Progress Matrix,Log"
    Src1Range = "A1:T60"
    Src2Row = 2
    srcCol = 1
    
    Set srcWb = ActiveWorkbook
    Set dstWb = Application.Workbooks.Open(PathName)
    'Set dstwb = Application.Workbooks("Progress.xlsx")
    
    PageNames = Split(PageNames, ",")
    
    srcWb.Activate
    
    srcWb.Worksheets(PageNames(0)).Range(Src1Range).Copy
    Application.DisplayAlerts = False
    
    With dstWb
        .Activate
        .Worksheets(PageNames(0)).Activate
        .ActiveSheet.Range("A1").Activate
        .ActiveSheet.Paste
        Selection.PasteSpecial Paste:=xlPasteColumnWidths
        Selection.PasteSpecial Paste:=xlPasteFormats
    End With
    Application.DisplayAlerts = True
    
    Call ClearClipboard
    
    'row scan
    With srcWb.Worksheets(PageNames(1))
        While .Cells(Src2Row, 1) <> ""
            'col scan
            While .Cells(2, srcCol) <> ""
                dstWb.Worksheets(PageNames(1)).Cells(Src2Row, srcCol).Value = .Cells(Src2Row, srcCol).Value
                srcCol = srcCol + 1
            Wend
            Src2Row = Src2Row + 1
            srcCol = 1
        Wend
    End With
    
    dstWb.Close savechanges:=True

End Sub


Private Sub logHeader()
' Purpose:
' Accepts:
' Returns:

'A Date
    Const szInt As Long = 255
    Dim lBuffer As Long
    
    Dim computerName As String * szInt
    Dim userName As String * szInt
    
    'timer reset goes here
    
    Worksheets(logSheet).Range("A" & rowEntry & ":M" & rowEntry).Insert Shift:=xlShiftDown
    
    Worksheets(logSheet).Range("A" & rowEntry).Value = Date
    
    'B Time
    Worksheets(logSheet).Range("B" & rowEntry).Value = Time()
    
    'C PC Name
    lBuffer = szInt
    GetComputerName computerName, lBuffer
    Worksheets(logSheet).Range("C" & rowEntry).Value = Trim(TrimNull(computerName))
    
    'D User ID
    lBuffer = szInt
    GetUserName userName, lBuffer
    Worksheets(logSheet).Range("D" & rowEntry).Value = Trim(TrimNull(userName))
    
    'E Path
    Worksheets(logSheet).Range("E" & rowEntry).Value = ActiveWorkbook.Path & "\" & ActiveWorkbook.name

End Sub
