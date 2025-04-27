VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_FTPbased_Email"
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
Event Action(pMsg As String)
Dim strFileToEmailPath As String
Dim strConnectionString As String
Dim strGSPrintLocation As String
Public WithEvents xT As XTimer
Attribute xT.VB_VarHelpID = -1
Dim strCurrentDocType As String
Dim zFTPServices As New z_FTP_Services
Dim cs As String
Dim Logging As New z_Logging
Dim oXML As New zXML
Dim strDocCode As String

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

Public Sub Action()
Dim strFilePathToSend As String
Dim err As Long
    
    RaiseEvent Action("Email: getting files")
    MsgWaitObj 800
    
    strFilePathToSend = GetFilePathToSend()
    If strFilePathToSend > "" Then
        If CreatePDF(strFilePathToSend, strCurrentDocType) Then
            If UploadPDF(oXML.PDFFilename_fp, fs.GetFileName(oXML.PDFFilename_fp), err) Then
                SendMail strFilePathToSend
            End If
        End If
    End If
    RaiseEvent Action("Idle")
End Sub
Public Sub EnableTimer(bOn As Boolean)
    On Error GoTo errHandler
    xT.Enabled = bOn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.EnableTimer(bOn)", bOn
End Sub
'''''Private Sub xT_Tick()
'''''    On Error GoTo errHandler
'''''Dim strFilePathToSend As String
'''''Dim err As Long
'''''
'''''    xT.Enabled = False
'''''    Forms(0).Enabled = False
'''''    RaiseEvent STATUS("Email: getting files")
'''''    MsgWaitObj 800
'''''
'''''    strFilePathToSend = GetFilePathToSend()
'''''    If strFilePathToSend > "" Then
'''''        If CreatePDF(strFilePathToSend, strCurrentDocType) Then
'''''            If UploadPDF(oXML.PDFFilename_fp, fs.GetFileName(oXML.PDFFilename_fp), err) Then
'''''                SendMail strFilePathToSend
'''''            End If
'''''        End If
'''''    End If
'''''    RaiseEvent STATUS("Idle")
'''''
'''''    xT.Enabled = True
'''''    Forms(0).Enabled = True
'''''
'''''    Exit Sub
'''''
'''''errHandler:
'''''    If ErrMustStop Then Debug.Assert False: Resume
'''''    ErrorIn "z_FTPbased_Email.xT_Tick"
'''''End Sub
Private Function CreatePDF(strFilePathToSend As String, strDocType As String) As Boolean
    On Error GoTo errHandler
Dim res As Boolean
Dim sTemp As String

    oXML.SetPDFName fs.GetBaseName(strFilePathToSend)
    CreatePDF = True
    Select Case strDocType
    Case "PO"
        sTemp = strSharedRootFolder & "\Templates\PO_DOC_FO.XSL"
    End Select

    oXML.SetFilePath strFilePathToSend, sTemp
    res = oXML.CreateFOUsingMSXML
        If res = False Then
            LogSaveToFile "PrintXML: res:" & CStr(res)
            CreatePDF = False
            Exit Function
        End If
    
    res = oXML.ConvertFOtoPDF
        If res = False Then
            LogSaveToFile "Dispatcher: zXML.PrintXML: call to ConvertFOtoPDF failed"
            CreatePDF = False
        End If
            

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.CreatePDF(strFilePathToSend,strDocType)", Array(strFilePathToSend, _
         strDocType)
End Function
Public Function UploadPDF(strFilePathToSend As String, DestinationFilename As String, pErr As Long) As Boolean
    On Error GoTo errHandler
Dim zFTPServices As New z_FTP_Services

    UploadPDF = zFTPServices.UploadFile("OWN", strFilePathToSend, DestinationFilename, pErr)
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.UploadPDF(strFilePathToSend,DestinationFilename)", _
         Array(strFilePathToSend, DestinationFilename)
End Function
Public Sub SendMail(strFilePathToSend As String)
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
 
    iAccepted = 0
    
    RaiseEvent STATUS("")
    Set poSendMail = New clsSendMail
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
    
    MailFile strFilePathToSend

    Dim sbasename As String
    sbasename = Left(strFilePathToSend, Len(strFilePathToSend) - Len(fs.GetExtensionName(strFilePathToSend)) - 1)
    If fs.FileExists(strFilePathToSend) Then
        fs.DeleteFile sbasename & ".*", True
    End If
    Set poSendMail = Nothing
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.SendMail"
End Sub

Private Sub MailFile(XMLFile_fp As String)
    On Error GoTo errHandler
    On Error GoTo errHandler
Dim i As Integer
Dim strLine As String
Dim strAddress As String
Dim strXMLFile As String
Dim strWholeMessage As String
Dim strReference As String
Dim XMLFile As ujXML
Dim strSender As String
Dim tmp As String
Dim res As Boolean

    If XMLFile_fp = "" Then Exit Sub
    bSendComplete = True
    Logging.Log "Attempting to send . . . " & fs.GetBaseName(XMLFile_fp)
    res = ParseXMLFile(XMLFile_fp, XMLFile, strWholeMessage, strAddress, strReference, strSender, strSendersEmail, strDocCode)
    
    If res Then
        oXML.SetPDFName fs.GetBaseName(XMLFile_fp)
        If oXML.CreateHTMLEmail(strSharedRootFolder & "\Templates\PO_LDOC_html.XSLT", strWholeMessage) Then
            If SendOneMessage(strWholeMessage, Replace(strFileToEmailPath, ".HTML", ".PDF"), _
                strDocCode, IIf(bTestMode, strEmailFrom, strAddress), strReference, strSender, strSendersEmail) Then
                Logging.Log fs.GetBaseName(XMLFile_fp) & " sent to: " & IIf(bTestMode, strEmailFrom, strAddress)
            End If
        End If
    Else
        Logging.Log "Cannot parse file" & fs.GetBaseName(XMLFile_fp)
    End If
    Set XMLFile = Nothing
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.MailFile"
End Sub
Private Function SendOneMessage(pContent As String, pAttachmentPath As String, pDocCode As String, pAddress As String, pReference As String, pSenderName As String, pSendersEmail As String) As Boolean
    On Error GoTo errHandler
    On Error GoTo errHandler
Dim lngCode As Long
Dim BytesSent As Long
Dim Nullstring As String
Dim Buffer As String * 81
Dim fs As FileSystemObject
Dim strErr As String

    poSendMail.Subject = pReference & ": " & strSubject
    
    strFailMsg = pDocCode & poSendMail.Subject & Time() & " Not sent to: " & pAddress
    strSuccessMsg = pDocCode & poSendMail.Subject & " Sent to: " & pAddress
    If pContent = "" Then pContent = "Please open attached document"
    poSendMail.Message = pContent                      ' Optional
    poSendMail.Recipient = pAddress                 ' Required, separate multiple entries with delimiter character
    poSendMail.RecipientDisplayName = pAddress      ' Optional, separate multiple entries with delimiter character
    poSendMail.FromDisplayName = pSenderName
    poSendMail.From = pSendersEmail

    Nullstring = Chr$(0)
    Set fs = New FileSystemObject
    bSendComplete = False
    poSendMail.send

    Set fs = Nothing
    bSendComplete = True
    SendOneMessage = True
    Exit Function
errHandler:
    ErrPreserve
    SendOneMessage = False
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.SendOneMessage(pContent,pAttachmentPath,pDocCode,pAddress,pReference," & _
  "pSenderName,pSendersEmail)", Array(pContent, pAttachmentPath, pDocCode, pAddress, pReference, _
   pSenderName, pSendersEmail)
End Function

Private Sub ClearFolder()
    On Error GoTo errHandler

    fs.DeleteFile strFileToEmailPath, True
Dim strFilename As String
Dim i As Integer
Dim fls
Dim fol
Dim f As File

    Set fol = fs.GetFolder(strSourceFolder)
    Set fls = fol.Files
    For Each f In fls
        f.Delete True
    Next

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.ClearFolder"
End Sub
Public Sub GetEmailSettings()
    On Error GoTo errHandler
    
    strSourceFolder = strSharedRootFolder & "\EMAILS_Links\"
    strSMTPServer = GetProperty("SMTPServer")
    strSMTP_Username = GetProperty("SMTP_Username")
    strSMTP_Password = GetProperty("SMTP_Password")
    strEmailFrom = GetProperty("EmailFrom")
    strSubject = GetProperty("Subject")
    strSenderName = GetProperty("SenderName")
    bTestMode = GetProperty("TestMode") = "TRUE"
    strGSPrintLocation = GetProperty("GSPRINTLocation")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.GetSettings"
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    
    Logging.SetLogFile oPC.SharedFolderRoot & "\Logs\EmailLog.txt"
    Set xT = New XTimer
    xT.Interval = 5000
    xT.Enabled = False
    LoadProperties
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.Class_Initialize"
End Sub
Public Sub SetTimerEnabled(Val As Boolean)
    xT.Enabled = Val
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
    ErrorIn "z_FTPbased_Email.GetProperty(pKey)", pKey
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
    
    ADOConn.open "Provider=SQLOLEDB.1;Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";NConnect Timeout=45"
    
    sSQL = "SELECT * FROM tProperty"
    Set rsProperty = New ADODB.Recordset
    rsProperty.CursorLocation = adUseClient
    rsProperty.open sSQL, ADOConn, adOpenKeyset, adLockOptimistic
    Set rsProperty.ActiveConnection = Nothing
    
    ADOConn.Close
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.LoadProperties", , , ADOConn
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
    ErrorIn "z_FTPbased_Email.NameOfPC"
End Property

''''''Public Property Get TimerEnabled() As Boolean
''''''    On Error GoTo errHandler
''''''    TimerEnabled = xT.Enabled
''''''    Exit Property
''''''errHandler:
''''''    If ErrMustStop Then Debug.Assert False: Resume
''''''    ErrorIn "z_FTPbased_Email.TimerEnabled"
''''''End Property

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
    ErrorIn "z_FTPbased_Email.poSendMail_SendFailed(explanation)", explanation
End Sub

Private Sub poSendMail_SendSuccesful()
    On Error GoTo errHandler
Dim strMsg As String

    bSendComplete = True
    iAccepted = iAccepted + 1
    RaiseEvent STATUS(strSuccessMsg & "   " & Format(Now, "dd/mm/yy Hh:Nn AM/PM"))
    UpdateLog "Emailed", strDocCode
    If fs.FileExists(PathAndBasename(strFileToEmailPath) & ".XML") Then
        fs.DeleteFile PathAndBasename(strFileToEmailPath) & ".*"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.poSendMail_SendSuccesful"
End Sub


Public Property Let SharedRootFolder(Val As String)
    On Error GoTo errHandler
    strSharedRootFolder = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.SharedRootFolder(Val)", Val
End Property
Public Property Let EmailFrom(Val As String)
    On Error GoTo errHandler
    strEmailFrom = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.EmailFrom(Val)", Val
End Property
'Public Property Let SourceFolder(Val As String)
'    strSourceFolder = Val
'End Property
Public Property Let Subject(Val As String)
    On Error GoTo errHandler
    strSubject = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.Subject(Val)", Val
End Property
Public Property Let SenderName(Val As String)
    On Error GoTo errHandler
    strSenderName = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.SenderName(Val)", Val
End Property
Public Property Let TestMode(Val As Boolean)
    On Error GoTo errHandler
    bTestMode = Val
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.TestMode(Val)", Val
End Property

Private Function ExtractDocCodeFromFilename(pIn As String) As String
    On Error GoTo errHandler
    ExtractDocCodeFromFilename = Right(fs.GetBaseName(pIn), Len(fs.GetBaseName(pIn)) - InStr(1, fs.GetBaseName(pIn), "_"))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.ExtractDocCodeFromFilename(pIn)", pIn
End Function
Private Function PathAndBasename(pIn As String) As String
    On Error GoTo errHandler
    PathAndBasename = fs.GetParentFolderName(pIn) & "\" & fs.GetBaseName(pIn)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.PathAndBasename(pIn)", pIn
End Function
Private Sub UpdateLog(pMsg As String, pDocCode As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    If ADOConn.State = 0 Then
        ADOConn.Provider = "sqloledb"
        ADOConn.open strMainConnectionString
    End If
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
    ErrorIn "z_FTPbased_Email.UpdateLog(pMsg,pDocCode)", Array(pMsg, pDocCode)
End Sub

Public Property Let ConnectionString(Val As String)
    cs = Val
End Property

Private Function GetFilePathToSend() As String
    On Error GoTo errHandler
Dim mFileName As String
Dim mFileToEmailPath As String

    mFileName = Dir(strSourceFolder & "*.XML", vbNormal)
    
    If mFileName > "" Then
        mFileToEmailPath = strSourceFolder & mFileName
    Else
        mFileToEmailPath = ""
    End If
    If InStr(1, mFileToEmailPath, "PO") > 0 Then
        strCurrentDocType = "PO"
    End If
    GetFilePathToSend = mFileToEmailPath
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_FTPbased_Email.GetFilePathToSend"
End Function

Private Function ParseXMLFile(File_fp As String, ujXMLDoc As ujXML, AccompanyingMsg As String, Destination As String, Reference As String, Sender As String, SendersEMail As String, DOCode As String) As Boolean
    On Error GoTo errHandler
    ParseXMLFile = False
    If fs.FileExists(File_fp) Then
        Set ujXMLDoc = New ujXML
        ujXMLDoc.docReadFromFile File_fp, "UNICODE"
        ujXMLDoc.navTop
        If ujXMLDoc.docXslLocate("DestinationAddress") Then
            Destination = ujXMLDoc.Element.Text
        Else
            Destination = strEmailFrom
        End If
        ujXMLDoc.navTop
        If ujXMLDoc.docXslLocate("AccompanyingMessage") Then
           If ujXMLDoc.Element.Text > "" Then
               AccompanyingMsg = ujXMLDoc.Element.Text
           End If
        End If
        ujXMLDoc.navTop
        If ujXMLDoc.docXslLocate("SendersEmail") Then
            SendersEMail = ujXMLDoc.Element.Text
        Else
            SendersEMail = strEmailFrom
        End If
        If strSendersEmail = "" Then strSendersEmail = strEmailFrom
        ujXMLDoc.navTop
        If ujXMLDoc.docXslLocate("StaffMember") Then
            Sender = ujXMLDoc.Element.Text
        Else
            Sender = strEmailFrom
        End If
        If Sender = "" Then Sender = SendersEMail
        ujXMLDoc.navTop
        If ujXMLDoc.docXslLocate("DocCode") Then
            DOCode = ujXMLDoc.Element.Text
        Else
            DOCode = "Unknown"
        End If
       ParseXMLFile = True
    End If
    Exit Function
errHandler:
    Exit Function
End Function

