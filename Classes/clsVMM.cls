VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private oDct As Object

Sub class_initialize()
    Set oDct = CreateObject("Scripting.Dictionary")
    
End Sub

Sub class_terminate()
    Set oDct = Nothing
    
End Sub

Property Let oIdxVal(ByVal index As String, ByVal value As Variant)
    
    Select Case TypeName(value)
    Case "Boolean", "Byte", "Integer", "Long", "Currency", "Single", "Double", "Date", "String"
        oDct(index) = value
    Case Else
        Set oDct(index) = value
    End Select
    
End Property

Property Set oIdxVal(ByVal index As String, ByVal value As Variant)
   
    Set oDct(index) = value
   
End Property

Property Get oIdxVal(ByVal index As String) As Variant

    Select Case TypeName(oDct(index))
    Case "Boolean", "Byte", "Integer", "Long", "Currency", "Single", "Double", "Date", "String"
        oIdxVal = oDct(index)
    Case Else
        Set oIdxVal = oDct(index)
    End Select
    
End Property
