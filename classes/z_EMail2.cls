VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_EMail"
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

Event STATUS(pMsg As String)

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
    
   ' RaiseEvent STATUS("COMPLETE")
    xT.Enabled = True
    Exit Sub

End Sub

Private Sub poSendMail_SendFailed(explanation As String)

    bSendComplete = True
    
    LogTransmission "", strFailMsg & "   " & explanation & "  " & Format(Now, "dd/mm/yy Hh:Nn AM/PM")
    
    RaiseEvent STATUS(strFailMsg & "   " & explanation & "  " & Format(Now, "dd/mm/yy Hh:Nn AM/PM"))
End Sub

Private Sub poSendMail_SendSuccesful()

    bSendComplete = True
    iAccepted = iAccepted + 1
    
    LogTransmission "", strSuccessMsg & "   " & Format(Now, "dd/mm/yy Hh:Nn AM/PM")
    RaiseEvent STATUS(strSuccessMsg & "   " & Format(Now, "dd/mm/yy Hh:Nn AM/PM"))

End Sub

Private Sub poSendMail_Status(STATUS As String)

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
 
 
    iAccepted = 0
    Nullstring = Chr$(0)
    NL = Chr$(13) + Chr$(10)
    GetFilesToPRint
    If UBound(arToMail) = 0 Then Exit Sub
    RaiseEvent STATUS("")
    Screen.MousePointer = vbHourglass
strErr = "pos 1"
 '''''   Set oS = New SimpleMail
    Set poSendMail = New clsSendMail
    With poSendMail

        ' **************************************************************************
        ' Optional properties for sending email, but these should be set first
        ' if you are going to use them
        ' **************************************************************************

        .SMTPHostValidation = VALIDATE_NONE         ' Optional, default = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX   ' Optional, default = VALIDATE_SYNTAX
        .Delimiter = ";"                            ' Optional, default = ";" (semicolon)

        ' **************************************************************************
        ' Basic properties for sending email
        ' **************************************************************************
        .SMTPHost = strSMTPServer               ' Required the fist time, optional thereafter
'        .From = strEmailFrom                       ' Required the fist time, optional thereafter
'        .FromDisplayName = strSenderName        ' Optional, saved after first use
'        .Recipient = txtTo.Text                     ' Required, separate multiple entries with delimiter character
'        .RecipientDisplayName = txtToName.Text      ' Optional, separate multiple entries with delimiter character
'        .CcRecipient = txtCc                        ' Optional, separate multiple entries with delimiter character
'        .CcDisplayName = txtCcName                  ' Optional, separate multiple entries with delimiter character
'        .BccRecipient = txtBcc                      ' Optional, separate multiple entries with delimiter character
'        .ReplyToAddress = strEmailFrom            ' Optional, used when different than 'From' address
'        .Subject = txtSubject.Text                  ' Optional
'        .Message = txtMsg.Text                      ' Optional
'        .Attachment = Trim(txtAttach.Text)          ' Optional, separate multiple entries with delimiter character

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
        .UserName = strSMTP_Username                     ' Optional, default = Null String
        .Password = strSMTP_Password                     ' Optional, default = Null String, value is NOT saved
        .POP3Host = ""
        .MaxRecipients = 3                        ' Optional, default = 100, recipient count before error is raised
        
        ' **************************************************************************
        ' Advanced Properties, change only if you have a good reason to do so.
        ' **************************************************************************
        ' .ConnectTimeout = 10                      ' Optional, default = 10
        ' .ConnectRetry = 5                         ' Optional, default = 5
        ' .MessageTimeout = 60                      ' Optional, default = 60
        ' .PersistentSettings = True                ' Optional, default = TRUE
        ' .SMTPPort = 25                            ' Optional, default = 25

 
    End With
    
    MailAllWaiting

    ClearFolder
    Set poSendMail = Nothing
    Screen.MousePointer = vbDefault
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_EMail.SendMail", , , , "strErr", Array(strErr)
End Sub


Private Function SendOneMessage(pContent As String, pAttachmentPath As String, pDocCode As String, pAddress As String, pReference As String, pSenderName As String, pSendersEmail As String) As Boolean
    On Error GoTo errHandler
Dim lngCode As Long
Dim BytesSent As Long
Dim Nullstring As String
Dim Buffer As String * 81
Dim fs As FileSystemObject
Dim strErr As String

    poSendMail.Subject = pReference & ": " & strSubject                 ' Optional
    poSendMail.Message = pContent                      ' Optional
    poSendMail.Recipient = pAddress                 ' Required, separate multiple entries with delimiter character
    poSendMail.RecipientDisplayName = "Recipient display name"      ' Optional, separate multiple entries with delimiter character
    poSendMail.FromDisplayName = pSenderName
    'MsgBox "In SendOneMessage: pSender name = " & pSenderName
    poSendMail.From = pSendersEmail
    'MsgBox "In SendOneMessage: strSendersEmail name = " & strSendersEmail
    

    strFailMsg = poSendMail.Subject & Time() & " Not sent: " & pAddress
    strSuccessMsg = poSendMail.Subject & Time() & " Sent: " & pAddress
    Nullstring = Chr$(0)
    Set fs = New FileSystemObject
    If fs.FileExists(pAttachmentPath) Then
      '  MsgBox "Found attachment"
        poSendMail.Attachment = pAttachmentPath         ' Optional, separate multiple entries with delimiter character
    Else
       ' MsgBox "Not found"
    End If
    bSendComplete = False
    poSendMail.send

    Set fs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_EMail.SendOneMessage(pContent,pAttachmentPath,pDocCode,pAddress,pReference)", _
         Array("text", pAttachmentPath, pDocCode, pAddress, pReference), , , "strErr", Array(strErr)
End Function


Private Sub ClearPrintedFiles()
Dim i As Integer
Dim fs As New FileSystemObject

On Error Resume Next

    i = 1
    Do While i <= UBound(arToMail)
        fs.DeleteFile arToMail(i)
        i = i + 1
    Loop
End Sub
Private Sub GetFilesToPRint()
Dim strFileName As String
Dim i As Integer
    ReDim arToMail(0)
    strFileName = Dir(strSourceFolder & "*.htm", vbNormal)
    i = 1
    
    Do While strFileName <> ""   ' Start the loop.
        ReDim Preserve arToMail(i)
        arToMail(i) = strSourceFolder & strFileName
        strFileName = Dir
        i = i + 1
    Loop
  '  MsgBox "X"
    MsgWaitObj 10000
  '  MsgBox "Z"
End Sub

Private Sub MailAllWaiting()
    On Error GoTo errHandler
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

'Open log
    If UBound(arToMail) = 0 Then Exit Sub
    Set oLog = New z_TextFile
    oLog.OpenTextFileToAppend strSharedServerFolder & "\EmailLog.txt"
    oLog.WriteToTextFile "Start of Set"
    oLog.WriteToTextFile "============"
    bSendComplete = True
    i = 1
    Do While i <= UBound(arToMail)
        strReference = Right(arToMail(i), Len(arToMail(i)) - InStr(1, arToMail(i), "O_") - 1)
        oTF.OpenTextFileToRead arToMail(i)
        strWholeMessage = oTF.ReadWholeFile
        oTF.CloseTextFile
        strXMLFile = Replace(arToMail(i), ".HTML", ".XML")
        If fs.FileExists(strXMLFile) Then
            Set xmlFile = New ujXML
            xmlFile.docReadFromFile strXMLFile, "UNICODE"
            xmlFile.navTop
            If xmlFile.docXslLocate("DestinationAddress") Then
                strAddress = xmlFile.Element.Text
            Else
                strAddress = strEmailFrom
            End If
            xmlFile.navTop
            If xmlFile.docXslLocate("SendersEmail") Then
           '     MsgBox "Found SendersEmail"
                strSendersEmail = xmlFile.Element.Text
            Else
            '    MsgBox "Cant find SendersEmail"
                strSendersEmail = strEmailFrom
            End If
            If strSendersEmail = "" Then strSendersEmail = strEmailFrom
            xmlFile.navTop
            If xmlFile.docXslLocate("StaffMember") Then
                strSender = xmlFile.Element.Text
            Else
                strSender = strEmailFrom
            End If
            If strSender = "" Then strSender = strEmailFrom
            xmlFile.navTop
            If xmlFile.docXslLocate("DocCode") Then
                strDocCode = xmlFile.Element.Text
            Else
                strDocCode = "Unknown"
            End If
            Set xmlFile = Nothing
            If bSendComplete Then
                If SendOneMessage(strWholeMessage, Replace(arToMail(i), ".HTML", ".PDF"), strDocCode, IIf(bTestMode, strEmailFrom, strAddress), strReference, strSender, strSendersEmail) Then
                    oLog.WriteToTextFile strAddress & ": " & strReference
                    fs.DeleteFile arToMail(i)
                    DoEvents
                End If
            End If
        End If
        i = i + 1
    Loop
    oLog.WriteToTextFile "End of Set" & vbCrLf
    oLog.CloseTextFile
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMailer.MailAllWaiting"
End Sub
Private Sub ClearFolder()
Dim strFileName As String
Dim i As Integer
Dim fs As FileSystemObject
Dim fls
Dim fol
Dim f As File

    RaiseEvent STATUS(vbCrLf & "Clearing files")
    
    Set fs = New FileSystemObject
    Set fol = fs.GetFolder(strSourceFolder)
    Set fls = fol.Files
    
    For Each f In fls
        f.Delete True
    Next
    
End Sub
Public Sub GetSettings()
'    strSMTPServer = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "SMTPServer", "")
'    strSMTP_Username = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "SMTP_Username", "")
'    strSMTP_Password = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "SMTP_Password", "")
'    strEmailFrom = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "EmailFrom", "")
'    strSubject = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "Subject", "")
'    strSenderName = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "SenderName", "")
'    bTestMode = (GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "TestMode", "TRUE") = "TRUE")
    
    strSourceFolder = strSharedRootFolder & "\EMAILS\"
    
    strSMTPServer = GetProperty("SMTPServer")
    strSMTP_Username = GetProperty("SMTP_Username")
    strSMTP_Password = GetProperty("SMTP_Password")
    strEmailFrom = GetProperty("EmailFrom")
    strSubject = GetProperty("Subject")
    strSenderName = GetProperty("SenderName")
    bTestMode = GetProperty("TestMode") = "TRUE"
    
End Sub

Private Sub Class_Initialize()
    Set xT = New XTimer
    xT.Interval = 5000
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
    If rsProperty.Fields.Count > 0 Then GetProperty = Trim(CStr(rsProperty.Fields(1)))
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
    ADOConn.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";Network Library=dbmssocn;Connect Timeout=45"
    
    sSQL = "SELECT * FROM tProperty"
    Set rsProperty = New ADODB.Recordset
    rsProperty.CursorLocation = adUseClient
    rsProperty.Open sSQL, ADOConn, adOpenKeyset, adLockOptimistic
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

