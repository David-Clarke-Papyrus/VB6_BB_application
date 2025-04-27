VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_EDI_Transmission"
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

Dim rs As ADODB.Recordset


Dim ADOConn As New ADODB.Connection
Dim rsProperty As New ADODB.Recordset
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Sub EnableTimer(bOn As Boolean)
    xT.Enabled = bOn
End Sub
Private Sub xT_Tick()
    On Error GoTo errHandler
    xT.Enabled = False
    
    SendFiles
    
    xT.Enabled = True
    Exit Sub

    Exit Sub
errHandler:
    ErrorIn "z_EDI_Transmission.xT_Tick"
End Sub
Public Property Get TimerEnabled() As Boolean
    TimerEnabled = xT.Enabled
End Property

Private Function ExtractDocCodeFromFilename(pIn As String) As String
    On Error GoTo errHandler
    ExtractDocCodeFromFilename = Right(fs.GetBaseName(pIn), Len(fs.GetBaseName(pIn)) - InStr(1, fs.GetBaseName(pIn), "_"))
    Exit Function
errHandler:
    ErrorIn "z_EDI_Transmission.ExtractDocCodeFromFilename(pIn)", pIn
End Function
Private Function PathAndBasename(pIn As String) As String
    On Error GoTo errHandler
    PathAndBasename = fs.GetParentFolderName(pIn) & "\" & fs.GetBaseName(pIn)
    Exit Function
errHandler:
    ErrorIn "z_EDI_Transmission.PathAndBasename(pIn)", pIn
End Function
Private Function ExtractDestinationFromFilename(pIn As String) As String
    On Error GoTo errHandler
    ExtractDestinationFromFilename = Left(fs.GetBaseName(pIn), InStr(1, fs.GetBaseName(pIn), "_") - 1)
    Exit Function
errHandler:
    ErrorIn "z_EDI_Transmission.ExtractDestinationFromFilename(pIn)", pIn
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
Private Sub GetFTPAddresss(SupplierAcno As String, FTP_Site As String, FTP_User As String, FTPPassword As String, FTPInitialFolder As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    If ADOConn.State = 0 Then ADOConn.open strMainConnectionString
    Set cmd = New ADODB.Command
    cmd.CommandText = "GetPOFTPAddress"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Acno", adVarChar, , 50, SupplierAcno)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FTPAddress", adVarChar, adParamOutput, 200)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FTPUser", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FTPPassword", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@FTPInitialFolder", adVarChar, adParamOutput, 50)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = ADOConn
    cmd.Execute
     FTP_Site = FNS(cmd.Parameters(1))
     FTP_User = FNS(cmd.Parameters(2))
     FTPPassword = FNS(cmd.Parameters(3))
     FTPInitialFolder = FNS(cmd.Parameters(4))
    Set cmd = Nothing

    ADOConn.Close

    Exit Sub
errHandler:
    ErrorIn "z_EDI_Transmission.GetFTPAddresss(SupplierAcno,FTP_Site,FTP_User,FTPPassword)", _
         Array(SupplierAcno, FTP_Site, FTP_User, FTPPassword)
End Sub
Private Sub InsertToTransmissionTable(Docno As String, sMode As String, FTP_Site As String, Success As String, pErrors As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    If ADOConn.State = 0 Then ADOConn.open strMainConnectionString
    Set cmd = New ADODB.Command
    cmd.CommandText = "InsertToTransmissionTable"
    cmd.CommandType = adCmdStoredProc
   
'InsertToTransmissionTable (@DOC VARCHAR(50),@DEST VARCHAR(200),@Mode VARCHAR(20),@RES VARCHAR(20))
   
    Set par = cmd.CreateParameter("@DOC", adVarChar, adParamInput, 50, Docno)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@DEST", adVarChar, adParamInput, 200, FTP_Site)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Mode", adVarChar, adParamInput, 20, sMode)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@RES", adVarChar, adParamInput, 20, Success)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Errors", adVarChar, adParamInput, 300, pErrors)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = ADOConn
    cmd.Execute
    Set cmd = Nothing

    ADOConn.Close

    Exit Sub
errHandler:
    ErrorIn "z_EDI_Transmission.InsertToTransmissionTable(Docno,sMode,FTP_Site,Success)", Array(Docno, _
         sMode, FTP_Site, Success)
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



Public Sub SendFiles()
10        On Error GoTo errHandler
      Dim lngResult As Long
      Dim Res As Boolean
      Dim Zip
      Dim FTP1 As New FTPClass
      Dim fs As New FileSystemObject
      Dim fc
      Dim FTPFile As FTPFileClass
      Dim strBUFolder As String
      Dim cmd As ADODB.Command
      Dim strPos As String
      Dim fol As folder
      Dim f As File
      Dim bSent As Boolean
          
          
20        Set fol = fs.GetFolder(strSharedServerFolder & "\" & FOL_EDI_PURCHASORDERS_SEND)
30        Set fc = fol.files
40        If fc.Count = 0 Then Exit Sub
50        bSent = False
60        For Each f In fc
70            SendFile f.Name, f.Path
80            f.Delete True
90            bSent = True
100       Next
110       If bSent Then RaiseEvent STATUS("Messages transmitted")
120   Exit Sub
130       Exit Sub
errHandler:
140       ErrorIn "z_EDI_Transmission.SendFiles"
End Sub

Private Sub SendFile(pFilename As String, pPath As String)
    On Error GoTo errHandler
''OPEN FTP Connection===========================================
Dim FTPAddress As String
Dim FTPUser As String
Dim FTPPassword As String
Dim FTPFolder As String
Dim FTP1 As New FTPClass
Dim bResult As Boolean
Dim sFilenameForDestination As String

        sFilenameForDestination = Right(fs.GetBaseName(pFilename), Len(fs.GetBaseName(pFilename)) - InStr(1, fs.GetBaseName(pFilename), "_")) & ".EDI"
    GetFTPAddresss ExtractDestinationFromFilename(pFilename), FTPAddress, FTPUser, FTPPassword, FTPFolder
    
    If FTPAddress > "" Then
        Check FTP1.OpenFTP(FTPAddress, FTPUser, FTPPassword, True), EXC_GENERAL, "Opening FTP"
        Check FTP1.SetCurrentFolder(FTPFolder), EXC_GENERAL, "Opening FTP"
        bResult = FTP1.PutFile(pPath, sFilenameForDestination, True)
        InsertToTransmissionTable sFilenameForDestination, "EDI using FTP", FTPAddress, IIf(bResult, "OK", "Failed"), Err.LastDllError
        FTP1.CloseFTP
    Else
        UpdateLog "Cant get FTP address for Acno - : " & ExtractDestinationFromFilename(pFilename), ExtractDocCodeFromFilename(pFilename)
        InsertToTransmissionTable fs.GetBaseName(pFilename), "EDI using FTP", FTPAddress, "Failed", "Cant get FTP address for Acno - : " & ExtractDestinationFromFilename(pFilename)
    End If

'''''''''''''''''''
   ' MsgBox FTPAddress & vbCrLf & FTPUser & vbCrLf & FTPPassword
    Exit Sub
errHandler:
    ErrorIn "z_EDI_Transmission.SendFile(pFilename,pPath)", Array(pFilename, pPath)
End Sub

Public Sub GetSettings()
    On Error GoTo errHandler
    
    strSourceFolder = strSharedRootFolder & "\EMAILS\"
    
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
    ErrorIn "z_EDI_Transmission.GetSettings"
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "PapyConn.GetProperty(pKey)", pKey
    Exit Function
errHandler:
    ErrorIn "z_EDI_Transmission.GetProperty(pKey)", pKey
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
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NameOfPC"
End Property


Public Function GetRecentTransmissions(Optional pSince As Date) As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    If ADOConn.State = 0 Then ADOConn.open strMainConnectionString
    Set cmd = New ADODB.Command
    cmd.CommandText = "GetRecentTransmissions"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Since", adDate, , , pSince)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = ADOConn
    
    rs.open cmd, , adOpenForwardOnly
    Set cmd = Nothing
    Set rs.ActiveConnection = Nothing
    ADOConn.Close
    
    Set GetRecentTransmissions = rs
    Exit Function
   
    Exit Function
errHandler:
    ErrorIn "z_EDI_Transmission.GetRecentTransmissions(pSince)", pSince
End Function



