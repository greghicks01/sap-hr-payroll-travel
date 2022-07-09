VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class Description
' =================
' Designed to handle connection to email
' Properties used to collect Recipients (To/CC)
' Subject line, body and signatire
' Attachments are just a list of files with paths
'
'    ...oooOOOooo...
' Usage
' =================
'Set ce = New clsEmail
'
'    ce.mailTo = c1 ' colection of email addresses
'    ce.mailCC = c2  ' colection of email addresses
'    ce.Subject = "<%test> Test"
'    ce.subjectReplace "<%test>", "Sep"
'    ce.Body = "body"
'    ce.sigBlock = "signed"
'    ce.sendEmail ' sends the email
'    ce = Nothing
'    ...oooOOOooo...

Option Explicit

Private cpvoSess            As Object
Private cpvoDB              As Object ' connection to lotus

Private cpvMailToCol        As Collection
Private cpvMailCCCol        As Collection
Private cpvAttachments      As Collection

Private cpvMailTo()         As String ' mail to list
Private cpvMailCC()         As String ' mail cc list
Private cpvMailTOCount      As Integer
Private cpvMailCCCount      As Integer

Private cpvMailSubject      As String
Private cpvMailBody         As String
Private cpvMailSignature    As String
Private cpvMailRecipients   As String
Private cpvTitle            As String

Private cpvResend           As Boolean  ' resend flag

Public Sub Class_Initialize()
' Purpose:
' Accepts:
' Returns:
    Dim subjectPage, subjectRng, bodyPage, bodyRng, sigPage, sigRng As String

    Set cpvoSess = CreateObject("Notes.NotesSession")
    Set cpvoDB = cpvoSess.GetDatabase("", "")
    Call cpvoDB.OPENMAIL

    Set cpvAttachments = New Collection
    
    ' init subject, body and sig?
    cpvTitle = "Subject"
    subjectPage = getSheet(cpvTitle)
    subjectRng = getRange(cpvTitle)
    cpvMailSubject = Worksheets(subjectPage).Range(subjectRng).value
    
    cpvTitle = "Body"
    bodyPage = getSheet(cpvTitle)
    bodyRng = getRange(cpvTitle)
    cpvMailBody = Join(Application.Transpose(Worksheets(bodyPage).Range(bodyRng)), vbCrLf)
    
    cpvTitle = "Sig Block"
    sigPage = getSheet(cpvTitle)
    sigRng = getRange(cpvTitle)
    cpvMailSignature = Join(Application.Transpose(Worksheets(sigPage).Range(sigRng)), vbCrLf)

    Exit Sub

MailNotOpen:
    MsgBox "Please start Lotus notes before sending a report", vbCritical And vbOKOnly, "Lotus Notes not running"
    Err.Raise Err.Number

End Sub

Public Sub Class_Terminate()
' Purpose:
' Accepts:
' Returns:
    Set cpvoDB = Nothing
    Set cpvoSess = Nothing

    Set cpvAttachments = Nothing
    
    Set cpvMailToCol = Nothing
    Set cpvMailCCCol = Nothing

End Sub

Public Sub bodyReplace(ByVal replaceMe As String, ByVal withMe As String)
' Purpose:
' Accepts:
' Returns:
    cpvMailBody = Replace(cpvMailBody, replaceMe, withMe)
    
End Sub

Public Sub subjectReplace(ByVal replaceMe As String, ByVal withMe As String)
' Purpose:
' Accepts:
' Returns:
    cpvMailSubject = Replace(cpvMailSubject, replaceMe, withMe)
    
End Sub

Public Property Let Resending(ByVal resend As Boolean)
' Purpose:
' Accepts:
' Returns:
    cpvResend = resend

End Property

Public Property Let Subject(ByVal subjectStr As String)
    cpvMailSubject = subjectStr

End Property

Public Property Let Body(ByVal bodyStr As String)
' Purpose:
' Accepts:
' Returns:
    cpvMailBody = bodyStr
    
End Property

Public Sub sigBlockReplace(ByVal replaceMe As String, ByVal withMe As String)
' Purpose:
' Accepts:
' Returns:
    cpvMailSignature = Replace(cpvMailSignature, replaceMe, withMe)

End Sub

Public Property Get mailTO() As Collection

    If cpvMailToCol Is Nothing Then Set cpvMailToCol = New Collection
    
    Set mailTO = cpvMailToCol
    
End Property

Public Property Let mailTO(c As Collection)
' Purpose: convert collection to array to mailto
' Accepts:
' Returns:
    Dim x As Integer
    
    If c Is Nothing Or c.Count = 0 Then Exit Property
    
    Set cpvMailToCol = c
    
    ReDim cpvMailTo(c.Count - 1)

    For x = 0 To c.Count - 1
        cpvMailTo(x) = c.Item(x + 1)
    Next
    
    cpvMailTOCount = c.Count

End Property

Public Property Get mailCC() As Collection

    If cpvMailCCCol Is Nothing Then Set cpvMailCCCol = New Collection
    
    Set mailCC = cpvMailCCCol
    
End Property

Public Property Let mailCC(c As Collection)
' Purpose: Collection to array for mailCC
' Accepts:
' Returns:
    Dim x As Integer
    
    If c Is Nothing Or c.Count = 0 Then Exit Property
   
    Set cpvMailCCCol = c
    
    ReDim cpvMailCC(c.Count - 1)
    
    For x = 0 To c.Count - 1
        cpvMailCC(x) = c.Item(x + 1)
    Next
    
    cpvMailCCCount = c.Count
    
End Property

Public Property Get attachments() As Collection
    
    Set attachments = cpvAttachments
    
End Property

Public Property Let attachments(c As Collection)
' Purpose:
' Accepts:
' Returns:
        Set cpvAttachments = c
    
End Property

Sub sendEmail()
' Purpose: Sends the email
' Accepts: NIL
' Returns: NIL
    Dim oItem As Object, _
        oDoc As Object
    Dim v As Variant
            
    Set oDoc = cpvoDB.CreateDocument
    Set oItem = oDoc.CreateRichTextItem("BODY")
    
    If cpvResend Then
        cpvMailSubject = "**UPDATED** " & cpvMailSubject
    End If
    
    With oDoc
        .Form = "Memo"
        .Subject = cpvMailSubject
        
        If cpvMailTOCount > 0 Then
            .SendTo = cpvMailTo
        End If
                
        If cpvMailCCCount > 0 Then
            .copyTo = cpvMailCC
        End If
        
        .Body = cpvMailBody & vbCrLf & cpvMailSignature
        .postdate = Date
        .SaveMessageOnSend = True
        
    End With
    
    ' attachments
    If cpvAttachments.Count > 0 Then
        For Each v In cpvAttachments
            oItem = oItem.EMBEDOBJECT(1454, "", v)
        Next
    End If
    
    oDoc.Visible = True
    oDoc.Send False
    
exit_SendAttachment:

    Set oDoc = Nothing
    Set oItem = Nothing
    
    Exit Sub
    
err_handler:
    ' use select case when possible
    If Err.Number = 7225 Then
        MsgBox "File doesn't exist"
    Else
        MsgBox Err.Number & " " & Err.Description
    End If
    
    On Error GoTo exit_SendAttachment
    
End Sub

