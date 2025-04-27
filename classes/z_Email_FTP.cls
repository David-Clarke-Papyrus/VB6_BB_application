VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Email_FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim strSharedRootFolder As String
Dim strSMTPServer As String
Dim strSMTP_Username As String
Dim strSMTP_Password As String
Dim strEmailFrom As String
Dim strSourceFolder As String
Dim strSubject As String
Dim strSenderName As String
Dim bTestMode As Boolean
Dim arToMail() As String
Dim iAccepted As Integer
Dim bSendComplete As Boolean
Dim fInet As wininet
Dim strFailMsg As String
Dim strSuccessMsg As String
Dim strSendersEmail As String
Dim fs As New FileSystemObject
Event STATUS(pMsg As String)
Dim strFileToEmailPath As String
Dim strConnectionString As String
Dim strGSPrintLocation As String
Public WithEvents xT As XTimer
Attribute xT.VB_VarHelpID = -1


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

Dim ADOConn As New ADODB.Connection
Dim rsProperty As New ADODB.Recordset
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Sub EnableTimer(bOn As Boolean)
    xT.Enabled = bOn
End Sub
Private Sub xT_Tick()
    xT.Enabled = False
    
    SendMail
    
    xT.Enabled = True
    
    Exit Sub

End Sub
Public Property Get TimerEnabled() As Boolean
    TimerEnabled = xT.Enabled
End Property

Private Sub poSendMail_SendFailed(explanation As String)
    On Error GoTo errHandler
Dim strDocCode As String

    bSendComplete = True
 ' MsgBox "Failed " & explanation
    RaiseEvent STATUS(strFailMsg & "   " & explanation & "  " & Format(Now, "dd/mm/yy Hh:Nn AM/PM"))
    strDocCode = Right(fs.GetBaseName(strFileToEmailPath), Len(fs.GetBaseName(strFileToEmailPath)) - InStr(1, fs.GetBaseName(strFileToEmailPath), "_"))
    UpdateLog "Email FAILED:" & explanation, strDocCode
    If fs.FileExists(PathAndBasename(strFileToEmailPath) & ".XML") Then
        fs.DeleteFile PathAndBasename(strFileToEmailPath) & ".*"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_EMail.poSendMail_SendFailed(explanation)", explanation
End Sub

Private Sub poSendMail_SendSuccesful()
    On Error GoTo errHandler
Dim strDocCode As String
Dim strMsg As String

    bSendComplete = True
  ' MsgBox "successful"
    iAccepted = iAccepted + 1
    RaiseEvent STATUS(strSuccessMsg & "   " & Format(Now, "dd/mm/yy Hh:Nn AM/PM"))
    strDocCode = ExtractDocCodeFromFilename(strFileToEmailPath)
    UpdateLog "Emailed", strDocCode
    If fs.FileExists(PathAndBasename(strFileToEmailPath) & ".XML") Then
        fs.DeleteFile PathAndBasename(strFileToEmailPath) & ".*"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_EMail.poSendMail_SendSuccesful"
End Sub
Private Function ExtractDocCodeFromFilename(pIn As String) As String
    ExtractDocCodeFromFilename = Right(fs.GetBaseName(pIn), Len(fs.GetBaseName(pIn)) - InStr(1, fs.GetBaseName(pIn), "_"))
End Function
Private Function PathAndBasename(pIn As String) As String
    PathAndBasename = fs.GetParentFolderName(pIn) & "\" & fs.GetBaseName(pIn)
End Function
Private Sub UpdateLog(pMsg As String, pDocCode As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    If ADOConn.State = 0 Then ADOConn.open strMainConnectionString
    Set cmd = New ADODB.Command
    cmd.CommandText = "UpdateTransmissionLog"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@DOCCODE", adVarChar, , 15, pDocCode)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Msg", adVarChar, adParamInput, 1000, pMsg)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = ADOConn
    cmd.Execute
    
    Set cmd = Nothing

    ADOConn.Close
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_EMail.UpdateLog(pMsg,pDocCode)", Array(pMsg, pDocCode)
End Sub
Private Sub poSendMail_Status(STATUS As String)

'MsgBox STATUS
''    ' vbSendMail 'Status Event'
''    lstStatus.AddItem Status
''    lstStatus.ListIndex = lstStatus.ListCount - 1
''    lstStatus.ListIndex = -1

End Sub


Public Property Let SharedRootFolder(Val As String)
    strSharedRootFolder = Val
End Property
Public Property Let EmailFrom(Val As String)
    strEmailFrom = Val
End Property
Public Property Let SourceFolder(Val As String)
    strSourceFolder = Val
End Property
Public Property Let Subject(Val As String)
    strSubject = Val
End Property
Public Property Let SenderName(Val As String)
    strSenderName = Val
End Property
Public Property Let TestMode(Val As Boolean)
    bTestMode = Val
End Property


Public Sub SendMail()
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
 '   Screen.MousePointer = vbHourglass
 
    iAccepted = 0
    Nullstring = Chr$(0)
    NL = Chr$(13) + Chr$(10)
    strFileToEmailPath = ""
    GetFileToPrint
    If strFileToEmailPath = "" Then Exit Sub
    RaiseEvent STATUS("")
    p 1
    Set poSendMail = New clsSendMail
    p 2
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
        .SMTPHost = strSMTPServer               ' Required the fist time, optional thereafter
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
        .enCodeType = MyEncodeType                  ' Optional, default = MIME_ENCODE
        .Priority = etPriority                      ' Optional, default = PRIORITY_NORMAL
        .Receipt = True                         ' Optional, default = FALSE
        .UseAuthentication = strSMTP_Username > ""             ' Optional, default = FALSE
        .UsePopAuthentication = bPopLogin           ' Optional, default = FALSE
        .Username = strSMTP_Username                     ' Optional, default = Null String
        .Password = strSMTP_Password                     ' Optional, default = Null String, value is NOT saved
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
    MailFile
    p 6
    Dim sbasename As String
    sbasename = Left(strFileToEmailPath, Len(strFileToEmailPath) - Len(fs.GetExtensionName(strFileToEmailPath)) - 1)
    If fs.FileExists(sbasename & ".*") Then fs.DeleteFile sbasename & ".*", True
  '  If fs.FileExists(sbasename & ".HTML") Then fs.DeleteFile sbasename & ".HTML"
  '  If fs.FileExists(sbasename & ".HTML") Then fs.DeleteFile sbasename & ".HTML"
    Set poSendMail = Nothing
   ' MsgBox "Before reset of pointer"
    Screen.MousePointer = vbDefault
    p 7
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_EMail.SendMail", , , , "strErr,strErrPos", Array(strErr, strErrPos)
End Sub


Private Function SendOneMessage(pContent As String, pAttachmentPath As String, pDocCode As String, pAddress As String, pReference As String, pSenderName As String, pSendersEmail As String) As Boolean
10        On Error GoTo errHandler
      Dim lngCode As Long
      Dim BytesSent As Long
      Dim Nullstring As String
      Dim buffer As String * 81
      Dim fs As FileSystemObject
      Dim strErr As String

         ' MsgBox "In SendOneMessage"

20        poSendMail.Subject = pReference & ": " & strSubject
          
30        strFailMsg = poSendMail.Subject & Time() & " Not sent: " & pAddress
40        strSuccessMsg = poSendMail.Subject & Time() & " Sent: " & pAddress
50        If pContent = "" Then pContent = "Please open attached document"
60        poSendMail.Message = pContent                      ' Optional
70        poSendMail.Recipient = pAddress                 ' Required, separate multiple entries with delimiter character
80        poSendMail.RecipientDisplayName = pAddress      ' Optional, separate multiple entries with delimiter character
90        poSendMail.FromDisplayName = pSenderName
100       poSendMail.From = pSendersEmail

110       Nullstring = Chr$(0)
120       Set fs = New FileSystemObject
130       If fs.FileExists(pAttachmentPath) Then
140           poSendMail.Attachment = pAttachmentPath         ' Optional, separate multiple entries with delimiter character
150       End If
160       bSendComplete = False
          '    MsgBox "Before poSendMail.send"

170       poSendMail.send

180       Set fs = Nothing
190       bSendComplete = True
200       Exit Function
errHandler:
210       If ErrMustStop Then Debug.Assert False: Resume
220       ErrorIn "z_EMail.SendOneMessage(pContent,pAttachmentPath,pDocCode,pAddress,pReference)", _
               Array("text", pAttachmentPath, pDocCode, pAddress, pReference), , , "strErr", Array(strErr)
End Function
Private Sub GetFileToPrint()
10        On Error GoTo errHandler
      Dim strFileName As String
      Dim strFileName_PDF As String
      Dim i As Integer

20        strFileToEmailPath = ""
30        strFileName = Dir(strSourceFolder & "*.PDF", vbNormal)   'The PDF file is last created of the set
          
40        If strFileName = "" Then Exit Sub
50        strFileName_PDF = strFileName
60        strFileName = Replace(strFileName, ".PDF", ".HTML")
70        strFileToEmailPath = strSourceFolder & strFileName
80        If Not fs.FileExists(strFileToEmailPath) Then
             ' LogSaveToFile strFileToEmailPath & " does not exist"
90            Exit Sub
100       End If
          
110       Exit Sub
errHandler:
120       If ErrMustStop Then Debug.Assert False: Resume
130       ErrorIn "z_EMail.GetFileToPrint", , , , "strSourceFolder,strFilename,strFileName_PDF,strFileToEmailPath", Array(strSourceFolder, strFileName, strFileName_PDF, strFileToEmailPath)
End Sub

Private Sub MailFile()
10        On Error GoTo errHandler
      Dim i As Integer
      Dim strLine As String
      Dim strAddress As String
      Dim strDocCode As String
      Dim strXMLFile As String
      Dim strWholeMessage As String
      Dim strReference As String
      Dim fs As New FileSystemObject
      Dim oTF As New z_TextFile
      Dim xmlFile As ujXML
      Dim strSender As String
      Dim oLog As z_TextFile
      Dim tmp As String

      'Open log
20        If strFileToEmailPath = "" Then Exit Sub
30    p 11
40        Set oLog = New z_TextFile
50        oLog.OpenTextFileToAppend strSharedServerFolder & "\EmailLog.txt"
60    p 12
70        oLog.WriteToTextFile "Sending . . . " & fs.GetBaseName(strFileToEmailPath)
80        bSendComplete = True
90    p 13
100       tmp = fs.GetBaseName(strFileToEmailPath)
110       strReference = Right(tmp, Len(tmp) - InStr(1, tmp, "_"))
120       If fs.FileExists(strFileToEmailPath) Then
130           oTF.OpenTextFileToRead strFileToEmailPath
140           strWholeMessage = oTF.ReadWholeFile
150           oTF.CloseTextFile
160       End If
          
170   p 14
180       strXMLFile = Replace(strFileToEmailPath, ".HTML", ".XML")
190        If fs.FileExists(strXMLFile) Then
200            Set xmlFile = New ujXML
210            xmlFile.docReadFromFile strXMLFile, "UNICODE"
220            xmlFile.navTop
230            If xmlFile.docXslLocate("DestinationAddress") Then
240                strAddress = xmlFile.Element.Text
250            Else
260                strAddress = strEmailFrom
270            End If
280            xmlFile.navTop
290            If xmlFile.docXslLocate("AccompanyingMessage") Then
300               If xmlFile.Element.Text > "" Then
310                   strWholeMessage = xmlFile.Element.Text
320               End If
330            End If
340            xmlFile.navTop
350            If xmlFile.docXslLocate("SendersEmail") Then
360                strSendersEmail = xmlFile.Element.Text
370            Else
380                strSendersEmail = strEmailFrom
390            End If
400            If strSendersEmail = "" Then strSendersEmail = strEmailFrom
410            xmlFile.navTop
420            If xmlFile.docXslLocate("StaffMember") Then
430                strSender = xmlFile.Element.Text
440            Else
450                strSender = strEmailFrom
460            End If
470            If strSender = "" Then strSender = strEmailFrom
480            xmlFile.navTop
490            If xmlFile.docXslLocate("DocCode") Then
500                strDocCode = xmlFile.Element.Text
510            Else
520                strDocCode = "Unknown"
530            End If
540            Set xmlFile = Nothing
550            If bSendComplete Then
560                If SendOneMessage(strWholeMessage, Replace(strFileToEmailPath, ".HTML", ".PDF"), _
                    strDocCode, IIf(bTestMode, strEmailFrom, strAddress), strReference, strSender, strSendersEmail) Then
570                    oLog.WriteToTextFile strAddress & ": " & strReference
580                 '   DoEvents
590                End If
600            End If
610        End If
620   p 15
630       oLog.WriteToTextFile "End of Set" & vbCrLf
640       oLog.CloseTextFile
          
650       Exit Sub
errHandler:
660       If ErrMustStop Then Debug.Assert False: Resume
670       ErrorIn "frmMailer.MailAllWaiting", , , , "strErrPos", Array(strErrPos)
End Sub
Private Sub ClearFolder()

    fs.DeleteFile strFileToEmailPath, True
Dim strFileName As String
Dim i As Integer
Dim fls
Dim fol
Dim f As File

    Set fol = fs.GetFolder(strSourceFolder)
    Set fls = fol.files
    For Each f In fls
        f.Delete True
    Next

End Sub
Public Sub GetSettings()
    
    strSourceFolder = strSharedRootFolder & "\EMAILS\"
    
    strSMTPServer = GetProperty("SMTPServer")
    strSMTP_Username = GetProperty("SMTP_Username")
    strSMTP_Password = GetProperty("SMTP_Password")
    strEmailFrom = GetProperty("EmailFrom")
    strSubject = GetProperty("Subject")
    strSenderName = GetProperty("SenderName")
    bTestMode = GetProperty("TestMode") = "TRUE"
    strGSPrintLocation = GetProperty("GSPRINTLocation")
End Sub

Private Sub Class_Initialize()
    Set xT = New XTimer
    xT.Interval = 5000
    xT.Enabled = True
    LoadProperties
End Sub
Public Function GetProperty(pKey As String) As String
    On Error GoTo errHandler
    rsProperty.MoveFirst
    rsProperty.Find "PropertyKey = '" & pKey & "'"
    If rsProperty.EOF Then
        GetProperty = ""
        Exit Function
    End If
    If rsProperty.fields.Count > 0 Then GetProperty = Trim(CStr(rsProperty.fields(1)))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.GetProperty(pKey)", pKey
End Function
Public Function LoadProperties() As Boolean
    On Error GoTo errHandler
Dim sSQL As String
Dim strPCName As String
Dim strServerName As String

    strPCName = Trim(Me.NameOfPC)
    
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
    End If
    
    strServerName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", strPCName)
    
    ADOConn.Provider = "sqloledb"
    ADOConn.open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";NConnect Timeout=45"
    
    sSQL = "SELECT * FROM tProperty"
    Set rsProperty = New ADODB.Recordset
    rsProperty.CursorLocation = adUseClient
    rsProperty.open sSQL, ADOConn, adOpenKeyset, adLockOptimistic
    Set rsProperty.ActiveConnection = Nothing
    
    ADOConn.Close
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.LoadProperties"
End Function
Public Property Get NameOfPC() As String
    On Error GoTo errHandler
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NameOfPC"
End Property





