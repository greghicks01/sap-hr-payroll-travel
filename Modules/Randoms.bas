Attribute VB_Name = "Randoms"
Function getRandomValue(minV As Long, maxV As Long) As Long
' Purpose:
' Accepts:
' Returns:

    Dim rangeV As Long
    rangeV = maxV - minV
    getRandomValue = CLng(Rnd * CSng(rangeV)) + minV
    
End Function

Function randomDate(minY As Integer, maxY As Integer, Optional minM As Integer = 0, Optional maxM As Integer = 0, Optional minD As Integer = 0, Optional maxD As Integer = 0) As String
' Purpose:
' Accepts:
' Returns:
'
' P: generates an excel date serial for use in SAP
' A: Min/Max values for Year, Month and Days in past
' R:

    randomDate = Format( _
                            CDate( _
                                getRandomValue( _
                                    CLng(DateSerial(Year(Now) - minY, Month(Now) - minM, Day(Now) - minD)), _
                                    CLng(DateSerial(Year(Now) - maxY, Month(Now) - maxM, Day(Now) - maxD))) _
                                  ), "dd.mm.yyyy" _
                        )

End Function

Sub fillMissingData()
' Purpose:
' Accepts:
' Returns:

    ' gen data -
    Dim wsMissingData As Worksheet
    Dim destRow As Long
    Dim rl As String
    Dim pl As String
    
    Set oDctCommonData = CreateObject("scripting.dictionary")
    
    Randomize
    ' for loop and read off a page on the sheet
    
    Dim cags As clsAGSgen
    Set cags = New clsAGSgen
    
    For Each wsMissingData In ActiveWorkbook.Worksheets
    
        If wsMissingData.Visible = True And locateHeader(wsMissingData, "exeID") <> 0 Then
        
            With wsMissingData
        
                destRow = 2
                
                While .Cells(destRow, locateHeader(wsMissingData, "Level")) <> ""  ' Valid Rows

                    If .Cells(destRow, locateHeader(wsMissingData, "XL_Code_Control")) = "" Then 'Can add people data

                        'call buildPerson(wsMissingData, destRow , sTag As String)
                        'AGS
                        If .Cells(destRow, locateHeader(wsMissingData, "AGS_Nos")).Value = "" Then
                            fillCellValue wsMissingData, destRow, "AGS_Nos", cags.NextAGS
                        End If

                        ' Last Name
                        fillCellValue wsMissingData, destRow, "Last_Name", Worksheets("names").Cells(getRandomValue(2, 444), 3) 'Last_Name
                        ' First Name and Gender
                        temp = getRandomValue(2, 517)
                        fillCellValue wsMissingData, destRow, "First_Name", Worksheets("names").Cells(temp, 2).Value
                        fillCellValue wsMissingData, destRow, "Gender", Worksheets("names").Cells(temp, 1).Value
                        'Preferred name
                        fillCellValue wsMissingData, destRow, "Pref_Name", wsMissingData.Cells(destRow, locateHeader(wsMissingData, "First_Name")).Value
                        ' DOB
                        fillCellValue wsMissingData, destRow, "Date_of_Birth", randomDate(19, 64)
                        
                        ' address
                        fillCellValue wsMissingData, destRow, "House_Num_Street", CStr(getRandomValue(10000, 50000)) & " " & Worksheets("Address").Cells(getRandomValue(2, 28), 1).Value
                        tmp = getRandomValue(2, 28)
                        fillCellValue wsMissingData, destRow, "Town_Suburb", Worksheets("Address").Cells(tmp, 2).Value
                        fillCellValue wsMissingData, destRow, "State", Worksheets("Address").Cells(tmp, 3).Value
                        fillCellValue wsMissingData, destRow, "Post_Code", Worksheets("Address").Cells(tmp, 4).Value

                        tmp = getRandomValue(2, 28)
                        fillCellValue wsMissingData, destRow, "House_Num_Street_2", CStr(getRandomValue(10000, 50000)) & " " & Worksheets("Address").Cells(getRandomValue(2, 28), 1).Value
                        fillCellValue wsMissingData, destRow, "Town_Suburb_2", Worksheets("Address").Cells(tmp, 2).Value
                        fillCellValue wsMissingData, destRow, "State_2", Worksheets("Address").Cells(tmp, 3).Value
                        fillCellValue wsMissingData, destRow, "Post_Code_2", Worksheets("Address").Cells(tmp, 4).Value
                
                        'LEAVE
                        Select Case .Cells(destRow, locateHeader(wsMissingData, "PS_Area")).Value
                        Case Is = "CL" ' Blue CL
                            rl = "RL"
                            pl = "PM"
                        Case Is = "MC" ' Green MC
                            rl = "RF"
                            pl = "PF"
                        Case Is = "HS" ' Red HS
                            rl = "RL"
                            pl = "PM"
                        End Select
                                       
                        'logon id        'logon id
                        If wsMissingData.Cells(destRow, locateHeader(wsMissingData, "Existing_User")) = "Y" Then
                            Dim sUID As String
                            sUID = Left(.Cells(destRow, locateHeader(wsMissingData, "PS_Area")), 1) & _
                                Right(.Cells(destRow, locateHeader(wsMissingData, "AGS_Nos")), 5)
                                
                            fillCellValue wsMissingData, destRow, "Logon_Id", sUID
                        End If
                        
                        'fillCellValue wsMissingData, destRow, "Logon_Id", "Z" & sUnit & Right(.Cells(destRow, locateHeader(wsMissingData, "AGS_Nos")), 5)
                        fillCellValue wsMissingData, destRow, "REC_Leave", rl
                        fillCellValue wsMissingData, destRow, "Per_Leave", pl
                        
                        'PayScale
                        fillCellValue wsMissingData, destRow, "PS_Group", .Cells(destRow, locateHeader(wsMissingData, "Level"))
                        
                        ' Columns 38-90 PA40 SU01 PA30
                        For x = 38 To 90
                            fillCellValue wsMissingData, destRow, Worksheets("Default Data").Cells(x + 1, "B"), ""
                        Next
                
                    End If
                    
                     For x = 1 To 37
                        fillCellValue wsMissingData, destRow, Worksheets("Default Data").Cells(x + 1, "B"), ""
                    Next
                                            
                    destRow = destRow + 1
            
                Wend
                
            End With
                
        End If
        
    Next wsMissingData
    
    Set cags = Nothing
    
End Sub

Sub fillCellValue(dstWs As Worksheet, dstRow As Long, columnHead As String, newValue As String)
' Purpose:  fills nominated row in nominated cell with new value
'           fillCellValues is controlled by Default Data sheet column "Force"
'           Force will replace existing values when Default Force = "Y"
'
' Accepts:  dstWs As Worksheet = the target sheet for updating
'           dstRow As Long = the row where the data will go
'           columnHead As String = the string "header" in Row 1
'           newValue As String = the new value to apply, if "", automatically uses default
'
' Returns: NIL

    Dim forceWks As Worksheet
    Set forceWks = Worksheets("Default Data")
    
    ' find the Force value
    colNum = locateHeader(dstWs, columnHead)
    iRelRow = colNum + 1
    
    With dstWs
        Select Case forceWks.Cells(iRelRow, "D")
        Case Is = "Y"
            ' if forcing value change with newVal blank - use default
            If newValue = "" Then
                .Cells(dstRow, colNum) = forceWks.Cells(iRelRow, "C") 'Range("B2").Row
             Else
                .Cells(dstRow, colNum) = newValue
             End If
             
        Case Else
            If .Cells(dstRow, colNum) = "" Then
                If newValue = "" Then
                    .Cells(dstRow, colNum) = forceWks.Cells(iRelRow, "C")
                Else
                    .Cells(dstRow, colNum) = newValue
                End If
            End If
            
        End Select
        
    End With
    
End Sub
