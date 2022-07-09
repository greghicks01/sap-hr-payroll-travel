VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufStructure 
   Caption         =   "SAP Org Structure Designer"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13350
   OleObjectBlob   =   "ufStructure.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cbCancelled As Boolean

Private Const sXMLConfigPath = "XML Structure Path"

Private Sub CommandButton1_Click()
' Purpose: reads XML and builds tree
' Accepts:
' Returns:

    Dim v As String
    Dim oXML As Object, oXMLNode As Object
    
    ' get identifier
    Select Case MsgBox("Yes for sheet" & vbCr & vbLf & "No for file", vbYesNoCancel, "Sheet or File?")
        Case vbYes
            If Not dispUF(v) Then
                Exit Sub
            End If
            v = getConfigData(sXMLConfigPath) & v & ".xml"
            
        Case vbNo
            ' file dialog
            v = selectFileSystemItem(msoFileDialogFilePicker, getConfigData(sXMLConfigPath) & "*.xml")
            
             If v = "-1" Then
                 Exit Sub
             End If
        
        Case vbCancel
            Exit Sub
            
    End Select
    
    ' open the relevant XML based on id
    Set oXML = CreateObject("Microsoft.XMLDOM")
    
    oXML.Load v
    
    Me.tvStruct.Nodes.Clear
    
    'Set tRoot = oXML.ChildNodes(1)
    For Each sAttrib In oXML.FirstChild.Attributes
        For Each oControl In mpData.Pages(0).Controls
            If oControl.name = sAttrib.name Then
                oControl.Value = sAttrib.Value
            End If
        Next
    Next
    
    'There is one node for each worksheet in the master
    For Each oXMLNode In oXML.ChildNodes
        'Prepare the initial values to kick off this process
          
        Call xmlChilds(oXMLNode, Me.tvStruct)
        
    Next

End Sub

Sub xmlChilds(oXMLNode As Object, oTVNode As Object)
' Purpose:
' Accepts:
' Returns:
    Dim oChild As Object, oTVnode2 As Object, ot As Object, sTag As String
    Dim oDic As Object
    Set oDic = CreateObject("scripting.dictionary")

    For Each oChild In oXMLNode.ChildNodes
    
    Set oTVnode2 = oTVNode
    
        sTag = "node:" & oChild.nodeName & "|"
    
        If oChild.Attributes.Length > 0 Then
            For Each sAttrib In oChild.Attributes
                sTag = sTag & sAttrib.name & ":" & oChild.getAttribute(sAttrib.name) & "|"
            Next
        End If
        
        sTag = Left(sTag, Len(sTag) - 1)

        arTag = Split(sTag, "|")
        For Each el In arTag
            aNode = Split(arTag(0), ":")
            oDic.Item(aNode(0)) = aNode(1)
        Next
        
        nodeText = ""
        
        Select Case aNode(1)
        Case "pos"
            iAttribCount = 0
            For Each el In arTag
                If iAttribCount = 0 Or iAttribCount = 1 Or iAttribCount = 5 Or iAttribCount = 6 Then
                    aTmp = Split(el, ":", 2)
                    nodeText = nodeText & aTmp(1) & ", "
                End If
                iAttribCount = iAttribCount + 1
            Next
        Case "org"
            For Each el In arTag
                aTmp = Split(el, ":", 2)
                nodeText = nodeText & aTmp(1) & ", "
            Next
        End Select
        
        nodeText = Mid(nodeText, 1, Len(nodeText) - 2)
    
        If Me.tvStruct.Nodes.Count = 0 Then
            Set ot = Me.tvStruct.Nodes.Add(, , "A" & CStr(Me.tvStruct.Nodes.Count + 1), nodeText)
        Else
            Set ot = Me.tvStruct.Nodes.Add(oTVnode2, tvwChild, "A" & CStr(Me.tvStruct.Nodes.Count + 1), nodeText)
        End If
        
        ot.Tag = sTag
        
        ' Recursive calls for each child below the current one
        Call xmlChilds(oChild, ot)
    Next
    
End Sub

Private Sub tvStruct_Click()
' Purpose:
' Accepts:
' Returns:

    Dim oDicTag As Object
    Set oDicTag = CreateObject("scripting.dictionary")
 
    If Me.tvStruct.Nodes.Count = 0 Then
        Exit Sub
    End If
    
    Set st = Me.tvStruct.Nodes(Me.tvStruct.SelectedItem.Key)
    
    If st.Tag = "" Then
        Exit Sub
    End If
    
    'Convert Tag to Data Dictionary
    arTag = Split(st.Tag, "|")
    For Each el In arTag
        aTmp = Split(el, ":")
        oDicTag.Item(aTmp(0)) = aTmp(1)
    Next
      
    ' 0         1
    'node:org|name:test1
    Select Case oDicTag.Item("node")
    Case "org"
        'set page to org and fill
        mpData.Value = 1
        oname = oDicTag.Item("name")
    Case "pos"
        'set page to pos and fill
        mpData.Value = 2
        'clear all the fields
        For Each oControl In mpData.SelectedItem.Controls
            If TypeName(oControl) = "TextBox" Or TypeName(oControl) = "ComboBox" Then
                oControl.Value = ""
            ElseIf TypeName(oControl) = "ListBox" Then
                oControl.Clear
            End If
        Next
        'fill it
        ' 0         1           2   3       4       5           6           7
        'node:pos|level:APS3|qty:6|fill:6|org:HS|position:False|name:test|roles:a;b;s
        For Each oControl In mpData.SelectedItem.Controls
                If TypeName(oControl) = "ListBox" Then
                    arList = Split(oDicTag.Item("roles"), ";")
                    For Each sItem In arList
                        oControl.AddItem sItem
                    Next
                Else
                    For Each sKey In oDicTag
                        If InStr(1, oControl.name, sKey) <> 0 Then
                            oControl.Value = oDicTag.Item(sKey)
                        End If
                    Next
                End If
        Next
        
    End Select
    
End Sub

Private Sub cbAdd_Click()
' Purpose: Adds a node to the users selected node
' Accepts:
' Returns:

    'load form then apply data
    Dim v As String
    Dim sNode As String, sPref As String
    v = ""
    sPref = "A"
  
    If Me.tvStruct.Nodes.Count > 0 Then
    
        If Me.tvStruct.SelectedItem Is Nothing Then Exit Sub
        
        st = Me.tvStruct.Nodes(Me.tvStruct.SelectedItem.Key)
        If InStr(1, st, ",") <> 0 Then
            v = Left(Trim(st), InStr(1, st, ",") - 1)
        Else
            v = st
        End If
    End If

    ' if user cancels
    If Not dispUFNC(v) Then
        Exit Sub
    End If
    
    ' 0         1           2   3       4       5           6           7
    'node:pos|level:APS3|qty:6|fill:6|org:HS|position:False|name:test|roles:a;b;s
    'node:org|name:test1
    oAr = Split(v, "|")
    Select Case Right(oAr(0), 3)
    Case "pos"
        '0,1,5 6
        For x = 0 To UBound(oAr)
            oAtmp = Split(oAr(x), ":")
            
            If x = 0 Or x = 1 Or x = 5 Or x = 6 Then
                sv = sv & oAtmp(1) & ", "
            End If
        Next
    Case "org"
        For x = 0 To UBound(oAr)
            oAtmp = Split(oAr(x), ":", 2)
            sv = sv & oAtmp(1) & ", "
        Next
    End Select

    If InStr(1, sv, ",") Then
        sv = Left(sv, Len(sv) - 2)
    End If
       
    If Me.tvStruct.Nodes.Count = 0 Then
        Set ot = Me.tvStruct.Nodes.Add(, , "A1", sv)
    Else
        On Error GoTo bad_node_id
        sNode = sPref & CStr(Me.tvStruct.Nodes.Count + 1)
        Set ot = Me.tvStruct.Nodes.Add(Me.tvStruct.SelectedItem.Key, tvwChild, sNode, sv)
        On Error GoTo 0
    End If
    ' Set tag with all data collected
    ot.Tag = v
    
    Exit Sub

bad_node_id:
    sNode = Chr(Asc(sPref) + 1) & CStr(Me.tvStruct.Nodes.Count + 1)
    Resume
    
End Sub

Private Sub cbCancel_Click()
' Purpose:
' Accepts:
' Returns:
    Me.Hide
End Sub

Private Sub cbRemove_Click()
' Purpose:
' Accepts:
' Returns:

    If Me.tvStruct.SelectedItem Is Nothing Then
        Exit Sub
    End If
    
    Me.tvStruct.Nodes.Remove Me.tvStruct.SelectedItem.Key
    
End Sub

Private Sub cbSave_Click()
' Purpose:
' Accepts:
' Returns:

    Dim oXML As Object
    Dim oXMLNode As Object
    Dim oTVNode As Object
        
    If Me.tvStruct.Nodes.Count = 0 Then
        Exit Sub
    End If
    
    Set oXML = CreateObject("Microsoft.XMLDOM")
    
    Set OnEWeL = oXML.createElement("root")
    For Each oControl In mpData.Pages(0).Controls
        If TypeName(oControl) = "TextBox" Then
            OnEWeL.setattribute(oControl.name) = oControl.Value
        End If
    Next
    oXML.appendChild OnEWeL
 
    smallTrav Me.tvStruct.Nodes(1), oXML, oXML.ChildNodes(0)
    
    oXML.Save (getConfigData(sXMLConfigPath) & cDatasetIdent.Value & ".xml")
    
End Sub
Sub smallTrav(n As Object, oXML As Object, oXMLNode As Object)
' Purpose:
' Accepts:
' Returns:

    Dim OnEWeL As Object
    
    Set objSiblingNode = n
    
    Do
        
        Set oDic = CreateObject("scripting.dictionary")
        ' process the treeview node for inclusion in XML
        oAr = Split(objSiblingNode.Tag, "|")
        
        'load the array into the Dic object
        For Each oItem In oAr
            oTmp = Split(oItem, ":")
            oDic.Item(oTmp(0)) = oTmp(1)
        Next
        
        ' Create the new XML node
        Set OnEWeL = oXML.createElement(oDic.Item("node"))
         
        'Remove these before the "Pos" test below
        oDic.Remove "node"
            
        ' load the remaining data as attributes if this is a pos node
        'If OnEWeL.nodename = "pos" Then
        For Each sKey In oDic
            OnEWeL.setattribute(sKey) = oDic.Item(sKey)
        Next
        'End If
        
        'add the node to XML DOM object
        oXMLNode.appendChild OnEWeL
        Set oDic = Nothing
        
        'Continue processing the treeview
        If Not objSiblingNode.Child Is Nothing Then
            Call smallTrav(objSiblingNode.Child, oXML, OnEWeL)
        End If
        Set objSiblingNode = objSiblingNode.Next
        
    Loop While Not objSiblingNode Is Nothing
    
End Sub
Private Sub cbPopSheet_Click()
' Purpose:
' Accepts:
' Returns:

    Dim dstWs As Worksheet, srcHdr As Worksheet, dstRow As Long
    
    If Me.tvStruct.Nodes.Count = 0 Then
        Exit Sub
    End If
       
    dstRow = 2
    ' find or create ws
    'If Worksheets(cDatasetIdent.Value) Is Nothing Then
    
        Set dstWs = ActiveWorkbook.Sheets.Add(After:=Sheets(ActiveWorkbook.Sheets.Count))
        ' add headers from AC Hire
        Set srcHdr = Sheets("ACHire")
        iCol = 1
        While srcHdr.Cells(1, iCol) <> ""
            dstWs.Cells(1, iCol) = srcHdr.Cells(1, iCol)
            dstWs.Cells(1, iCol).Interior.Color = srcHdr.Cells(1, iCol).Interior.Color
            dstWs.Cells(1, iCol).Columns.AutoFit
            iCol = iCol + 1
        Wend
        dstWs.Range("A1").AutoFilter
    
     
        dstWs.Cells(dstRow, locateHeader(dstWs, "exeID")) = cDatasetIdent.Value
        dstWs.name = Replace(cDatasetIdent.Value, " ", "_")
        
    'Else
        'dstWs = Worksheets(cDatasetIdent.Value)
        
    'End If
    
    XLBuildTraverse Me.tvStruct.Nodes(1), dstRow, dstWs

End Sub

Sub XLBuildTraverse(n As Object, destRow As Long, dstWs As Worksheet)
' Purpose:
' Accepts:
' Returns:

    'Dim oDic As Object
    
    Randomize

    Set objSiblingNode = n
    
    Do
        Set oDic = CreateObject("scripting.dictionary")
        ' process the treeview node for inclusion in XML
            oAr = Split(objSiblingNode.Tag, "|")
            
        ' 0         1           2   3       4       5           6          7
        ' node:pos|level:APS3|qty:6|fill:6|org:HS|position:False|name:test|roles:a;d;b
        ' node:org|name:test1

        ' load the array into the Dic object
        For Each oItem In oAr
            oTmp = Split(oItem, ":")
            oDic.Item(oTmp(0)) = oTmp(1)
        Next
                 
        ' load the remaining data as attributes if this is a pos node
        Select Case oDic.Item("node")
        Case "pos"
            ' for loop - counting the qty/fill options
            For x = 1 To CInt(oDic.Item("qty"))
                'build pos
                Call prepNewSystem(dstWs, destRow)
                Call buildPosition(dstWs, destRow, objSiblingNode.Tag)
                    
                If x <= CInt(oDic.Item("fill")) Then
                    'build person
                    Call buildPerson(dstWs, destRow, objSiblingNode.Tag)
                Else
                    dstWs.Cells(destRow, locateHeader(dstWs, "XL_Code_Control")) = "N"
                End If
                destRow = destRow + 1
            Next
            
        Case "org"
            ' build org
            Call buildOrg(dstWs, destRow, objSiblingNode.Tag)
            
        End Select
        
        Set oDic = Nothing
        
        'Continue processing the treeview
        If Not objSiblingNode.Child Is Nothing Then
            Call XLBuildTraverse(objSiblingNode.Child, destRow, dstWs)
        End If
        Set objSiblingNode = objSiblingNode.Next
        
    Loop While Not objSiblingNode Is Nothing

End Sub

Private Sub UserForm_Initialize()
' Purpose: Clears out the tree
' Accepts: NIL
' Returns: NIL

    Me.tvStruct.Nodes.Clear
    Set oNameCounter = New clsNameGen
    
End Sub

Private Sub UserForm_Terminate()
' Purpose: Termination of the form cleans up new objects
' Accepts: NIL
' Returns: NIL

    Set oNameCounter = Nothing
End Sub
