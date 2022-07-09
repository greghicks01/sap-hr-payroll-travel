Attribute VB_Name = "DataLoad"


Sub prepNewSystem(dstWs As Worksheet, dstRow As Long)
' Purpose: Prepares a new system by updating values based on col 1 to 6 in Default
' Accepts: dstWs destination worksheet
' Returns: nIL
' Columns 1 - 6
    Dim forceWks As Worksheet
    Dim oDct As Object
    Set forceWks = Worksheets("Default Data")
    For x = 1 To 6
        fillCellValue dstWs, dstRow, forceWks.Cells(x + 1, "B"), ""
    Next


End Sub

Sub buildOrg(dstWs As Worksheet, ByVal destRow As Long, sTag As String)
' Purpose: builds org data where required, cols 7-20
' Accepts:
' Returns:
' Columns 7-20 PP03Org
    Dim forceWks As Worksheet, sTmpVal As String
    Dim oDct As Object
    Set oDctData = CreateObject("scripting.dictionary")
    
    ' 0         1 sTag
    'node:org|name:test1
    arData = Split(sTag, "|")
    For Each el In arData
        aTmp = Split(el, ":")
        oDctData.Item(aTmp(0)) = aTmp(1)
    Next
    
    
    Set forceWks = Worksheets("Default Data")
    For x = 7 To 20
        fillCellValue dstWs, destRow, forceWks.Cells(x + 1, "B"), ""
    Next
    
    sTmpVal = oDctData.Item("name")
    'Org_Name with [n..]
    If InStr(1, oDctData.Item("name"), "[n") > 0 Then
    
        arTmpName = Split(oDctData.Item("name"), "[")
        arTmpName(0) = Trim(arTmpName(0))
        arTmpName(1) = Mid(arTmpName(1), 1, InStr(1, arTmpName(0), "]") - 1)
        
        oNameCounter.oDct.Item(arTmpName(0)) = PadNumber(CInt(oNameCounter.oDct.Item(arTmpName(0))) + 1, Len(arTmpName(1)))
        sTmpVal = oNameCounter.oDct.Item(arTmpName(0))
    
    End If
    
    fillCellValue dstWs, destRow, "Org_Unit_Name", sTmpVal

End Sub

Sub buildPosition(dstWs As Worksheet, ByVal destRow As Long, sTag As String)
' Purpose: build a position into the data cols 21-37
' Accepts:
' Returns:
' columns 21-37 PP03Pos
    
    Dim forceWks As Worksheet, sTmpVal As String
    Dim oDct As Object
    Set oDctData = CreateObject("scripting.dictionary")
    
    ' sTag
    ' 0         1           2   3       4       5           6           7
    'node:pos|level:APS3|qty:6|fill:6|org:HS|position:False|name:test|roles:zhr;zhe
    arData = Split(sTag, "|")
    For Each el In arData
        aTmp = Split(el, ":")
        oDctData.Item(aTmp(0)) = aTmp(1)
    Next
    
    Set forceWks = Worksheets("Default Data")
    For x = 21 To 37
        fillCellValue dstWs, destRow, forceWks.Cells(x + 1, "B"), ""
    Next
    
    If InStr(1, oDctData.Item("level"), "APS") > 0 Or InStr(1, oDctData.Item("level"), "EL") > 0 Then
        fillCellValue dstWs, destRow, "Level", oDctData.Item("level")
        
    ElseIf InStr(1, oDctData.Item("level"), "SEC") > 0 Then
        fillCellValue dstWs, destRow, "Level", "DHS-SEC"
        fillCellValue dstWs, destRow, "ESG_for_CAP", "5"
        
    ElseIf InStr(1, oDctData.Item("level"), "CEO") > 0 Then
        fillCellValue dstWs, destRow, "Level", "CEO"
        fillCellValue dstWs, destRow, "ESG_for_CAP", "5"
        
    Else 'must be SES band  more to do
        sTmp = Right(sLevel, 1)
        fillCellValue dstWs, destRow, "ESG_for_CAP", "5"
    End If
    
    'Pos_Name with [n..]
    sTmpVal = oDctData.Item("name")
    If InStr(1, oDctData.Item("name"), "[n") > 0 Then
    
        arTmpName = Split(oDctData.Item("name"), "[")
        arTmpName(0) = Trim(arTmpName(0))
        arTmpName(1) = Mid(arTmpName(1), 1, InStr(1, arTmpName(1), "]") - 1)
        
        oNameCounter.oDct.Item(arTmpName(0)) = PadNumber(CInt(oNameCounter.oDct.Item(arTmpName(0))) + 1, Len(arTmpName(1)))
        sTmpVal = arTmpName(0) + oNameCounter.oDct.Item(arTmpName(0))
    
    End If
    
    fillCellValue dstWs, destRow, "Plan_Version", ""
    fillCellValue dstWs, destRow, "Planning_Status", ""
    fillCellValue dstWs, destRow, "Start_Date", ""
    fillCellValue dstWs, destRow, "End_Date", ""
    fillCellValue dstWs, destRow, "Cost_Centre", ""
    fillCellValue dstWs, destRow, "Building_code", ""
    fillCellValue dstWs, destRow, "Pos_Name", sTmpVal
    fillCellValue dstWs, destRow, "Level", oDctData.Item("level")
    fillCellValue dstWs, destRow, "PS_Area", oDctData.Item("org")
    
    fillCellValue dstWs, destRow, "PS_Group", oDctData.Item("level")
    fillCellValue dstWs, destRow, "Pers_Area", ufStructure.cPers_area
    fillCellValue dstWs, destRow, "Pers_Sub", ufStructure.cPersSubArea
    fillCellValue dstWs, destRow, "DT_PP13_Roles", oDctData.Item("roles")
        
End Sub

Sub buildPerson(dstWs As Worksheet, ByVal destRow As Long, sTag As String)
' Purpose:
' Accepts:
' Returns:
' Columns 38-90 PA40 SU01 PA30
 'AGS
    Dim sUID As String
    Dim cags As clsAGSgen
    Set cags = New clsAGSgen
 
    With dstWs
    
        fillCellValue dstWs, destRow, "Existing_User", ""
        
        If .Cells(destRow, locateHeader(dstWs, "AGS_Nos")).Value = "" Then ' AND not a Non-Payroll Emplyee
            fillCellValue dstWs, destRow, "AGS_Nos", cags.NextAGS
        End If
        
        Set cags = Nothing
        
        fillCellValue dstWs, destRow, "Employee_Group", ""
        fillCellValue dstWs, destRow, "Employee_Subgroup", ""
        
        ' Last Name
        fillCellValue dstWs, destRow, "Last_Name", Worksheets("names").Cells(getRandomValue(2, 444), 3) 'Last_Name
        ' First Name and Gender
        temp = getRandomValue(2, 517)
        fillCellValue dstWs, destRow, "First_Name", Worksheets("names").Cells(temp, 2).Value
        fillCellValue dstWs, destRow, "Gender", Worksheets("names").Cells(temp, 1).Value
        'Preferred name
        fillCellValue dstWs, destRow, "Pref_Name", dstWs.Cells(destRow, locateHeader(dstWs, "First_Name")).Value
        ' DOB
        fillCellValue dstWs, destRow, "Date_of_Birth", randomDate(19, 64)
        
        'Nationality
        fillCellValue dstWs, destRow, "Nationality", ""
        
        fillCellValue dstWs, destRow, "Payroll", ufStructure.cPayScaleArea
        
        ' address
        fillCellValue dstWs, destRow, "House_Num_Street", CStr(getRandomValue(10000, 50000)) & " " & Worksheets("Address").Cells(getRandomValue(2, 28), 1).Value
        
        tmp = getRandomValue(2, 28)
        fillCellValue dstWs, destRow, "Town_Suburb", Worksheets("Address").Cells(tmp, 2).Value
        fillCellValue dstWs, destRow, "State", Worksheets("Address").Cells(tmp, 3).Value
        fillCellValue dstWs, destRow, "Post_Code", Worksheets("Address").Cells(tmp, 4).Value
        
        tmp = getRandomValue(2, 28)
        fillCellValue dstWs, destRow, "House_Num_Street_2", CStr(getRandomValue(10000, 50000)) & " " & Worksheets("Address").Cells(tmp, 1).Value
        fillCellValue dstWs, destRow, "Town_Suburb_2", Worksheets("Address").Cells(tmp, 2).Value
        fillCellValue dstWs, destRow, "State_2", Worksheets("Address").Cells(tmp, 3).Value
        fillCellValue dstWs, destRow, "Post_Code_2", Worksheets("Address").Cells(tmp, 4).Value
        
                          
        'LEAVE
        Select Case .Cells(destRow, locateHeader(dstWs, "PS_Area")).Value
        Case Is = "CL" ' Blue CL
            rl = "RL"
            pl = "PM"
            sUnit = "C"
        Case Is = "MC" ' Green MC
            rl = "RF"
            pl = "PF"
            sUnit = "M"
        Case Is = "HS" ' Red HS
            rl = "RL"
            pl = "PM"
            sUnit = "H"
        End Select
        
        'logon id
        If dstWs.Cells(destRow, locateHeader(dstWs, "Existing_User")) = "Y" Then
            sUID = Left(.Cells(destRow, locateHeader(dstWs, "PS_Area")), 1) & _
                sUnit & Right(.Cells(destRow, locateHeader(dstWs, "AGS_Nos")), 5)
            fillCellValue dstWs, destRow, "Logon_Id", sUID
        End If
        
        fillCellValue dstWs, destRow, "REC_Leave", CStr(rl)
        fillCellValue dstWs, destRow, "Per_Leave", CStr(pl)
        
        fillCellValue dstWs, destRow, "Password", ""
        
            Set forceWks = Worksheets("Default Data")
    For x = 38 To 90
        fillCellValue dstWs, destRow, forceWks.Cells(x + 1, "B"), ""
    Next

    
    End With
  
End Sub
