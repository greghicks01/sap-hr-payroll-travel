Attribute VB_Name = "TEst_Code"

Sub testlevel()
' Purpose:
' Accepts:
' Returns:
    Debug.Print locateHeader(ActiveSheet, "john")
    Debug.Print locateHeader(ActiveSheet, "Payroll")
    Debug.Print locateHeader(ActiveSheet, "Org_Unit_Name")
    Debug.Print locateHeader(ActiveSheet, "Last_Name")
    
End Sub
' Purpose:
' Accepts:
' Returns:

Sub tryRows()
' Purpose:
' Accepts:
' Returns:
    topRow = Range("B2").Row
    RelRow = CLng(Application.WorksheetFunction.Match("PP03_Org_i1002_Free_Text", Sheets("Default Data").Range("B2", "B200"), False)) - 1
    Debug.Print topRow
    Debug.Print RelRow
    Debug.Print Sheets("Default Data").Cells(topRow + RelRow, "C")
End Sub


Sub testAGS1()
    Dim a As clsAGSgen
    Set a = New clsAGSgen
    
    Debug.Print a.NextAGS
End Sub

Sub testAGS()
' Purpose:
' Accepts:
' Returns:
    Dim oAGS As clsAGSgen
    Set oAGS = New clsAGSgen
    
    Debug.Print oAGS.NextAGS
    
End Sub

Sub addCommentTabs()
' Purpose:
' Accepts:
' Returns:
    unhideAll
    hideSpecial
    Dim w As Worksheet
    
    For Each w In Worksheets
        With w
            If .Visible = True Then
                .Cells(2, 2).AddComment CStr(.Value)
            End If
        End With
    Next
End Sub

Sub findAllPersNsub()
' Purpose:
' Accepts:
' Returns:

    Dim oDct As Object, wb As Workbook, oFSO As Object, oFile As Object, w As Worksheet
    
    Set oDct = CreateObject("Scripting.Dictionary")

    For Each w In ActiveWorkbook.Worksheets
        If locateHeader(w, "exeID") <> 0 Then
            'scan Against Level
            ' get each Pers/Sud Area Unique
            tmpLvl = locateHeader(w, "Level")
            tmpPers = locateHeader(w, "Pers_Area")
            tmpPersSb = locateHeader(w, "Pers_Sub")
            sRow = 2
            With w
                While .Cells(sRow, tmpLvl) <> ""
                    oDct.Item(.Cells(sRow, tmpPers) & "," & .Cells(sRow, tmpPersSb)) = ""
                    sRow = sRow + 1
                Wend
            End With
        End If
    Next
    Set oFSO = CreateObject("Scripting.Filesystemobject")
    Set oFile = oFSO.CreateTextFile("S:\Automation\SAPQTP\QC Project\GUI\Data\pers.dat", True)
    
    For Each sKey In oDct
        oFile.Writeline sKey
    Next
End Sub

Sub updateCodes()
' Purpose:
' Accepts:
' Returns:
Dim ws As Worksheet

    For Each ws In ActiveWorkbook.Sheets
        srcRow = 2
        If ws.Visible = True Then
            With ws
                While .Cells(srcRow, locateHeader(ws, "Level")) <> ""
                    If .Cells(srcRow, locateHeader(ws, "PP03_i1005_Time_Unit")) <> "" Then
                        tmp = .Cells(srcRow, locateHeader(ws, "Activity_Group"))
                        .Cells(srcRow, locateHeader(ws, "Activity_Group")) = tmp & ";" & .Cells(srcRow, locateHeader(ws, "PP03_i1005_Time_Unit"))
                        .Cells(srcRow, locateHeader(ws, "PP03_i1005_Time_Unit")) = ""
                    End If
                    srcRow = srcRow + 1
                Wend
            End With
        End If
    Next
End Sub

Sub testPL()
' Purpose:
' Accepts:
' Returns:
Dim v As String
    
    v = "pos"
    
    If Not dispUFPickRole(v) Then
        Exit Sub
    End If
    
    'nodeType:pos|level:SEC|qty:1|fill:0|org:HS|position:True
    Debug.Print v
End Sub

Sub testUFNC()
' Purpose:
' Accepts:
' Returns:
    Dim v As String
    
    v = "pos"
    
    If Not dispUFNC(v) Then
        Exit Sub
    End If
    'cbNodeType:pos|cbLevel:SEC|tbQty:1|tbFill:0|cbOrg:HS|ckPosition:True
    Debug.Print v

End Sub

Sub testExp()
' Purpose:
' Accepts:
' Returns:
Dim v As String

    If Not dispUFExp(v) Then
        MsgBox "Cancelled"
        Exit Sub
    End If

    MsgBox v
    
End Sub

Sub testNodeConf()
' Purpose:
' Accepts:
' Returns:

    Load ufNodeConfig
    ufNodeConfig.tbParentType.Value = "pos"
    
    ufNodeConfig.Show
    
    Unload ufNodeConfig
    
    Load ufNodeConfig
    ufNodeConfig.tbParentType.Value = "org"
    
    ufNodeConfig.Show
    
    Unload ufNodeConfig

End Sub

Sub testdispUF()
' Purpose:
' Accepts:
' Returns:
Dim v As String

    If Not dispUF(v) Then
        MsgBox "user ignored or cancelled"
        Exit Sub
    End If
    
    MsgBox "body of code got " & v
End Sub

Sub testPrepNew()
' Purpose:
' Accepts:
' Returns:
    Dim dstWs As Worksheet
    prepNewSystem dstWs
End Sub


Sub testAGS2()
' Purpose:
' Accepts:
' Returns:

    Dim cags As clsAGSgen
    Set cags = New clsAGSgen
    
    For x = 1 To 20
        t = StrReverse(cags.NextAGS)
        s = 0
        
        For P = 1 To 8
            s = s + CInt(Mid(t, P, 1)) * 2 ^ P
        Next
        
        s = s Mod 11
        
        Debug.Print StrReverse(t) & " " & s
        
    Next
    
    Set cags = Nothing

End Sub

Sub testColumnDetect()

    IntColr = ActiveSheet.Cells(1, Selection.column).Interior.Color
    ActiveSheet.Cells(1, Selection.column).Interior.Color = xlColorIndexNone
    ActiveSheet.Cells(1, Selection.column).Interior.Color = IntColr

End Sub

Sub testColorPick()

    ufColourPicker.Show
    
    If ufColourPicker.bCancelled = True Then
        Exit Sub
    End If
    
    ActiveSheet.Cells(1, "F").Interior.Color = ufColourPicker.lgColourValue
    
End Sub


Sub testSet()
' Purpose:
' Accepts:
' Returns:

    setConfigData "RRRS", "Folder", "C:\Data"
    
End Sub
