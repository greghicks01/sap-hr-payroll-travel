Attribute VB_Name = "ProjectRibbon"
' add all the original buttons and other functionality here


Sub Btn_Export_QTP(Control As IRibbonControl)
' Purpose: exports for QTP to use in an execution
' Accepts:
' Returns:

    'Do not change the name of this macro, as it is hard-coded to your custom button.
    'This macro is tied to the 'To QTP' button in the 'Export' group under the SAP tab.

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    QTPTestPrep
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub
Sub Btn_Export_Payroll(Control As IRibbonControl)
' Purpose: Exports "payroll" on a tab by tab basis to top8fs1
' Accepts:
' Returns:

    'Do not change the name of this macro, as it is hard-coded to your custom button.
    'This macro is tied to the 'Payroll to Test Team' button in the 'Export' group under the SAP tab.

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    ExportPayroll
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub Btn_Update_Column_Restructure(Control As IRibbonControl)
' Purpose: exports to a temporary file to enable entire sheet column adjustment
' Accepts:
' Returns:

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    export_column_change
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
        
    Load ufColumnRestructure
    
    ufColumnRestructure.Show
    
    Unload ufColumnRestructure
    
End Sub

Sub Btn_Import_QTP(Control As IRibbonControl)
' Purpose: imports results from QTP.
' Accepts:
' Returns:

    'Do not change the name of this macro, as it is hard-coded to your custom button.
    'This macro is tied to the 'QTP Results' button in the 'Import' group under the SAP tab.

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    copyResultsFromDefault
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
 
End Sub

Sub Btn_Import_Pool(Control As IRibbonControl)
' Purpose:
' Accepts:
' Returns:

     'Do not change the name of this macro, as it is hard-coded to your custom button.
     'This macro is tied to the 'Pool Data' button in the 'Import' group under the SAP tab.
    
     'MsgBox "Replace this Message Box with a Call to your own Btn_Import_Pool or Procedure."
     Load ufStructure
     
     ufStructure.Show False
     
     'Unload ufStructure

End Sub

Sub Btn_Update_Header(Control As IRibbonControl)
' Purpose:
' Accepts:
' Returns:

    tmp = locateHeader(ActiveSheet, "exeID")
    
    If tmp = 0 Then Exit Sub
    
    payroll = ActiveSheet.Cells(2, locateHeader(ActiveSheet, "exeID")).Value 'B   Payroll
    
    unhideAll
    hideSpecial
    
    copyACHireHeadings
    
    UnHideHidden (payroll)

End Sub

Sub Btn_Update_genPerson(Control As IRibbonControl)
' Purpose:
' Accepts:
' Returns:

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    ' config Default
    fillMissingData
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub Btn_Update_Roles_Clean(Control As IRibbonControl)
' Purpose:
' Accepts:
' Returns:

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    scanAndCleanRoles
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub Btn_Update_NewSystem(Control As IRibbonControl) ' TODO: FIX REFERENCES
' Purpose: PREPARES FOR EXECUTION IN NEW SYSTEM
' Accepts:
' Returns:

    Dim ws As Worksheet, srcRow As Long
    Dim oC As Object
    
    Set oC = CreateObject("scripting.dictionary")
    'Set oUpdateList = CreateObject("scripting.dictionary")
    'Set oExceptList = CreateObject("scripting.dictionary")
    'fillCellValue dstWs,dstRow,colHead,newVal
    oC.Item("Parent") = "Y"
    oC.Item("Org_Unit_No.") = "Y"
    oC.Item("Position") = "Y"
    oC.Item("Sup_pos_no.") = "Y"
    oC.Item("Email") = "Y"
    oC.Item("Done") = "Y"
    oC.Item("Tax_Scale") = "Y"
    oC.Item("Bank_Details") = "Y"
    'oC.item("Existing_User" ?
    'oC.Item("Start_Date") = "Y"
    'PP03_Org_Object_Type
    'PP03_Org_BZOT_Office_Type
    'PP03_Org_i1002_Free_Text

        
    iDDRow = 2
    Do While Worksheets("Default Data").Cells(iDDRow, "B") <> ""
        sKey = Worksheets("Default Data").Cells(iDDRow, "B")
        If oC.exists(sKey) Then
            Worksheets("Default Data").Cells(iDDRow, "C") = ""
            Worksheets("Default Data").Cells(iDDRow, "D") = oC.Item(sKey)
        ElseIf sKey = "Start_Date" Then
            Worksheets("Default Data").Cells(iDDRow, "D") = "Y"
        ElseIf sKey = "PP03_Org_Object_Type" Or sKey = "PP03_Org_BZOT_Office_Type" Or sKey = "PP03_Org_i1002_Free_Text" Then
            Worksheets("Default Data").Cells(iDDRow, "C") = ""
        ElseIf 2 <= iDDRow And iDDRow <= 7 Or sKey = "Activity_Group" Then
            ' ignore these columns
        Else
            Worksheets("Default Data").Cells(iDDRow, "D") = ""
        End If
        iDDRow = iDDRow + 1
    Loop
    
    'unhideAll
    'hideSpecial

    For Each ws In Sheets

        srcRow = 2
        
        With ws
            If .Visible = True And locateHeader(ws, "exeID") <> 0 Then

                colLvl = locateHeader(ws, "Level")
                colRole = locateHeader(ws, "Activity_Group")
                'ROW scan
                While .Cells(srcRow, colLvl) <> ""
                
                    If InStr(1, .Cells(srcRow, colRole), "~") <> 0 Then
                        'remove client details
                        arActivityData = Split(.Cells(srcRow, colRole), "~")
                        .Cells(srcRow, colRole) = arActivityData(1)
                    End If
                    
                    For x = 1 To 37 ' find a way to read max value... there are exceptions above field 38
                        fillCellValue ws, srcRow, Worksheets("Default Data").Cells(x + 1, "B"), ""
                    Next

                    srcRow = srcRow + 1
                Wend
            End If
        End With
    Next

End Sub

Sub Btn_View_Payroll(Control As IRibbonControl)
' Purpose:
' Accepts:
' Returns:

    'Do not change the name of this macro, as it is hard-coded to your custom button.
    'This macro is tied to the 'Payroll' button in the 'View' group under the SAP tab.
    Dim payroll As String

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    
    If dispUF(payroll) Then
        UnHideHidden payroll
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

Sub Btn_View_Tab(Control As IRibbonControl)
' Purpose:
' Accepts:
' Returns:

    'Do not change the name of this macro, as it is hard-coded to your custom button.
    'This macro is tied to the 'Tab' button in the 'View' group under the SAP tab.
    ufTabList.Show

End Sub

Sub Btn_View_All(Control As IRibbonControl)
' Purpose:
' Accepts:
' Returns:

 'Do not change the name of this macro, as it is hard-coded to your custom button.
 'This macro is tied to the 'All' button in the 'View' group under the SAP tab.

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    unhideAll
    hideSpecial
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub
