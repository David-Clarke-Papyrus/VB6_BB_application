VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oS As SimpleMail

Public Sub InitializeMail()
    On Error GoTo errHandler
    Set oS = New SimpleMail
    oS.MailServer = oPC.SMTPServer
    oS.Sender = oPC.AdministratorEmail
    oS.SenderName = "NielsenSales"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Email.InitializeMail"
End Sub

Public Function SendMailA(pContent As String, pSubject As String, pDestination As String) As Boolean
Dim lngResult As Long

    On Error GoTo errHandler
'    oS.CustomHeaders = "Disposition-Notification-To: " & strEmailFrom
    oS.Subject = pSubject
    oS.MessageText = pContent
    oS.Recipient = pDestination
    lngResult = oS.SendEx
    Check (lngResult = 0), EXC_GENERAL, "Error calling Simplemail method SendEx: " & oS.LastStatusMessage

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Email.SendMail(pContent,pSubject,pDestination)", Array(pContent, pSubject, pDestination)
End Function
Public Function SendMailwithAttachment(pContent As String, pSubject As String, pDestination As String, pAttachment As String, pAttachmentName As String) As Boolean
Dim lngResult As Long

    On Error GoTo errHandler
'    oS.CustomHeaders = "Disposition-Notification-To: " & strEmailFrom
    oS.AttachFile pAttachment, pAttachmentName
    oS.Subject = pSubject
    oS.MessageText = pContent
    oS.Recipient = pDestination
    lngResult = oS.SendEx
    If lngResult <> 0 Then
        Err.Raise ERR_GENERAL, "SendMailwithAttachment", "Error calling Simplemail method SendEx: " & oS.LastStatusMessage
    End If
  '  Check (lngResult = 0), EXC_GENERAL, "Error calling Simplemail method SendEx: " & oS.LastStatusMessage

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Email.SendMail(pContent,pSubject,pDestination)", Array(pContent, pSubject, pDestination)
End Function

Public Sub TerminateMail()
    Set oS = Nothing
End Sub

