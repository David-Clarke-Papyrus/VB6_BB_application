VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_HOEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' *****************************************************************************
' Required declaration of the vbSendMail component (withevents is optional)
' You also need a reference to the vbSendMail component in the Project References
' *****************************************************************************
Private WithEvents poSendMail As vbSendMail.clsSendMail
Attribute poSendMail.VB_VarHelpID = -1

' misc local vars
Dim bAuthLogin      As Boolean
Dim bPopLogin       As Boolean
Dim bHtml           As Boolean
Dim MyEncodeType    As ENCODE_METHOD
Dim etPriority      As MAIL_PRIORITY
Dim bReceipt        As Boolean
Event STATUS(pMsg As String)

Dim rsConfig As New ADODB.Recordset
Dim sAttachmentPath As String
Dim fs As FileSystemObject
Dim strFailMsg As String
Dim strSuccessMsg As String

Private Sub Class_Initialize()
    Set poSendMail = New clsSendMail
End Sub

Private Sub Class_Terminate()
    Set poSendMail = Nothing
End Sub

Private Sub poSendMail_SendFailed(explanation As String)
    On Error GoTo errHandler
Dim strDocCode As String

    Set fs = New FileSystemObject
    RaiseEvent STATUS("Transmission has failed" & "   " & explanation & "  " & Format(Now, "dd/mm/yy Hh:Nn AM/PM"))
    strDocCode = Right(fs.GetBaseName(sAttachmentPath), Len(fs.GetBaseName(sAttachmentPath)) - InStr(1, fs.GetBaseName(sAttachmentPath), "_"))
    UpdateLog sAttachmentPath
    If fs.FileExists(PathAndBasename(sAttachmentPath) & ".HTML") Then
        fs.DeleteFile PathAndBasename(sAttachmentPath) & ".*"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_HOEMail.poSendMail_SendFailed(explanation)", explanation
End Sub

Private Sub poSendMail_SendSuccesful()
    On Error GoTo errHandler
Dim strDocCode As String
Dim strMsg As String
Dim fs As New FileSystemObject

    RaiseEvent STATUS("transmission successful" & "   " & Format(Now, "dd/mm/yy Hh:Nn AM/PM"))
    strDocCode = fs.GetBaseName(sAttachmentPath)
    UpdateLog "Failed: " & sAttachmentPath
  '  If fs.FileExists(PathAndBasename(sAttachmentPath) & ".HTML") Then
  '      fs.DeleteFile PathAndBasename(sAttachmentPath) & ".*"
  '  End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_HOEMail.poSendMail_SendSuccesful"
End Sub
Private Function ExtractDocCodeFromFilename(pIn As String) As String
    ExtractDocCodeFromFilename = fs.GetBaseName(pIn)
End Function

Private Function PathAndBasename(pIn As String) As String
    PathAndBasename = fs.GetParentFolderName(pIn) & "\" & fs.GetBaseName(pIn)
End Function
Private Sub UpdateLog(pFilename As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "UpdateTransmissionLog"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FILENAME", adVarChar, , 150, pFilename)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    
    Set cmd = Nothing

    Exit Sub

'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_HOEMail.UpdateLog(pMsg,pDocCode)", Array(pMsg, pDocCode)
    Exit Sub
errHandler:
    ErrorIn "z_HOEmail.UpdateLog(pFilename)", pFilename
End Sub

Private Sub poSendMail_Status(STATUS As String)

'MsgBox STATUS
''    ' vbSendMail 'Status Event'
''    lstStatus.AddItem Status
''    lstStatus.ListIndex = lstStatus.ListCount - 1
''    lstStatus.ListIndex = -1

End Sub

Public Sub PrepareSendMail()
    On Error GoTo errHandler
Dim lngCode As Long
Dim Length As Long
Dim BytesSent As Long
Dim Nullstring As String
Dim Temp As String * 256
Dim TheFile As String
Dim MSG As String
Dim NL As String
Dim X As String
Dim iResult As Integer
Dim strErr As String
Dim errPos As String
    Screen.MousePointer = vbHourglass
 
    Nullstring = Chr$(0)
    NL = Chr$(13) + Chr$(10)
    sAttachmentPath = ""
    RaiseEvent STATUS("")
    p 1
    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)
        p 3
        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = oPC.EMAIL_SMTP_Server               ' Required the fist time, optional thereafter
'        .From = strEmailFrom                       ' Required the fist time, optional thereafter
'        .FromDisplayName = strSenderName        ' Optional, saved after first use
'        .Recipient = txtTo.Text                     ' Required, separate multiple entries with delimiter character
 '       .RecipientDisplayName = txtToName.Text      ' Optional, separate multiple entries with delimiter character
'        .CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
'        .CcDisplayName = txtCcName                  ' Optional, separate multiple entries with delimiter character
'        .BccRecipient = txtBcc                      ' Optional, separate multiple entries with delimiter character
'        .ReplyToAddress = strEmailFrom            ' Optional, used when different than 'From' address
'        .Subject = txtSubject.Text                  ' Optional
'        .Message = txtMsg.Text                      ' Optional
'        .Attachment = Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character
        p 4
        ' **************************************************************************
        ' Additional Optional properties, use as required by your application / environment
        ' **************************************************************************
        .AsHTML = True                             ' Optional, default = FALSE, send mail as html or plain text
        .ContentBase = ""                           ' Optional, default = Null String, reference base for embedded links
        .EncodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = True                         ' Optional, default = FALSE
        .UseAuthentication = oPC.EMAIL_SMTP_Username > ""             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .Username = oPC.EMAIL_SMTP_Username                     ' Optional, default = Null String
        .Password = oPC.EMAIL_SMTP_Password                     ' Optional, default = Null String, value is NOT saved
        .POP3Host = ""
        .MaxRecipients = 3                        ' Optional, default = 100, recipient count before error is raised
        p 5
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        ' .SMTPPort = 25                            ' Optional, default = 25

 
    End With
    
    Screen.MousePointer = vbDefault
    p 7
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_HOEMail.SendMail", , , , "strErr,strErrPos", Array(strErr, strErrPos)
End Sub

Public Function SendOneMessage(pSubject As String, pContent As String, pAttachmentPath As String, pDocCode As String, pAddress As String, pReference As String, pSenderName As String, pSendersEmail As String) As Boolean
    On Error GoTo errHandler
Dim lngCode As Long
Dim BytesSent As Long
Dim Nullstring As String
Dim Buffer As String * 81
Dim fs As FileSystemObject
Dim strErr As String

    poSendMail.Subject = pSubject
    
    strFailMsg = poSendMail.Subject & Time() & " Not sent: " & pAddress
    strSuccessMsg = poSendMail.Subject & Time() & " Sent: " & pAddress
    
    poSendMail.Message = pContent                      ' Optional
    poSendMail.Recipient = pAddress                 ' Required, separate multiple entries with delimiter character
    poSendMail.RecipientDisplayName = pAddress      ' Optional, separate multiple entries with delimiter character
    poSendMail.FromDisplayName = pSenderName
    poSendMail.From = pSendersEmail

    Nullstring = Chr$(0)
    Set fs = New FileSystemObject
    If fs.FileExists(pAttachmentPath) Then
        poSendMail.Attachment = pAttachmentPath         ' Optional, separate multiple entries with delimiter character
        sAttachmentPath = pAttachmentPath
    End If
    poSendMail.send

    Set fs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_HOEMail.SendOneMessage(pContent,pAttachmentPath,pDocCode,pAddress,pReference)", _
         Array("text", pAttachmentPath, pDocCode, pAddress, pReference), , , "strErr", Array(strErr)
End Function

