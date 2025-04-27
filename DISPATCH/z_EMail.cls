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
'Mailing through SimpleMail
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
Dim oS As SimpleMail
Dim fInet As wininet
Dim strMsg As String

Event STATUS(pMsg As String)

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
    RaiseEvent STATUS("")
    Screen.MousePointer = vbHourglass
strErr = "pos 1"
    Set oS = New SimpleMail
strErr = "pos 2"
    oS.MailServer = strSMTPServer
    If strSMTP_Username > "" Then
        oS.UseAuthentication = 1
        oS.UserName = strSMTP_Username
        oS.Password = strSMTP_Password
    End If
strErr = "pos 3"
    oS.Sender = strEmailFrom
strErr = "pos 4"
    oS.SenderName = strSenderName
strErr = "pos 5"
    
strErr = "pos 6"
    GetFilesToPRint
    
strErr = "pos 7"
    MailAllWaiting

strErr = "pos 8"
    ClearFolder
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_EMail.SendMail", , , , "strErr", Array(strErr)
End Sub


Private Function SendOneMessage(pContent As String, pAttachmentPath As String, pDocCode As String, pAddress As String, pReference As String) As Boolean
    On Error GoTo errHandler
Dim lngCode As Long
Dim BytesSent As Long
Dim Nullstring As String
Dim Buffer As String * 81
Dim fs As FileSystemObject
Dim strErr As String

    Nullstring = Chr$(0)
strErr = "Pos 0"
    strMsg = ""
    oS.CustomHeaders = "Disposition-Notification-To: " & strEmailFrom
    oS.Subject = pReference & ": " & strSubject
strErr = "Pos 1"
    Set fs = New FileSystemObject
    strErr = "Pos 1o"
    If fs.FileExists(pAttachmentPath) Then
strErr = "Pos 1a"
        oS.ResetAttachments
strErr = "Pos 1b"
        oS.AttachFileWithContentType pAttachmentPath, "Order " & pDocCode & ".PDF", "application/pdf"
    End If
strErr = "Pos 2"
    oS.MessageHTMLText = pContent
    oS.Recipient = pAddress
strErr = "Pos 3"
    lngCode = oS.SendEx
strErr = "Pos 4"
    If lngCode = 0 Then 'OK
strErr = "Pos 4a"
        strMsg = strMsg & pDocCode & " Sent: " & pAddress & vbCrLf
        iAccepted = iAccepted + 1
        SendOneMessage = True
        LogTransmission pDocCode, "Transmitted:" & Format(Now, "dd/mm/yy Hh:Nn AM/PM")
    Else
strErr = "Pos 4b"
        strMsg = strMsg & Time() & " Not sent: " & pAddress & "(" & oS.LastStatusMessage & ")" & vbCrLf
        SendOneMessage = False
        LogTransmission pDocCode, "NOT TRANSMITTED:" & Format(Now, "dd/mm/yy Hh:Nn AM/PM")
    End If
strErr = "Pos 5"
    RaiseEvent STATUS(strMsg)
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
Dim strFilename As String
Dim i As Integer
    ReDim arToMail(0)
    strFilename = Dir(strSourceFolder & "*.htm", vbNormal)
    i = 1
    
    Do While strFilename <> ""   ' Start the loop.
        ReDim Preserve arToMail(i)
        arToMail(i) = strSourceFolder & strFilename
        strFilename = Dir
        i = i + 1
    Loop

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
Dim oLog As z_TextFile

'Open log
    Set oLog = New z_TextFile
    oLog.OpenTextFileToAppend strSharedServerFolder & "\EmailLog.txt"
    oLog.WriteToTextFile "Start of Set"
    oLog.WriteToTextFile "============"
    
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
            If xmlFile.docXslLocate("DocCode") Then
                strDocCode = xmlFile.Element.Text
            Else
                strDocCode = "Unknown"
            End If
            Set xmlFile = Nothing
            
            If SendOneMessage(strWholeMessage, Replace(arToMail(i), ".HTML", ".PDF"), strDocCode, IIf(bTestMode, strEmailFrom, strAddress), strReference) Then
                oLog.WriteToTextFile strAddress & ": " & strReference
                fs.DeleteFile arToMail(i)
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
Dim strFilename As String
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
    strSMTPServer = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "SMTPServer", "")
    'strSMTP_Username
    strSMTP_Username = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "SMTP_Username", "")
    strSMTP_Password = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "SMTP_Password", "")
    strEmailFrom = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "EmailFrom", "")
    strSourceFolder = strSharedRootFolder & "\EMAILS\"
    strSubject = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "Subject", "")
    strSenderName = GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "SenderName", "")
    bTestMode = (GetIniKeyValue(strSharedRootFolder & "\PBKS.INI", "EMAIL", "TestMode", "TRUE") = "TRUE")
End Sub
