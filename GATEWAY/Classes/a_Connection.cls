VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_connection"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private dteTimeStarted As Date
Private strPapyConnErr As String
'Private strServer As String
Private strServerName As String
Dim strPBKSSERVERMACHINE As String
Dim strLocalRootFolder As String
Dim strAdministratorEmail As String
Dim strSMTPServer As String
Dim strSMTPUsername As String
Dim strSMTPPassword As String
Dim strClientCode As String
Dim strSharedFolderRoot As String
Dim strCENTRALFTPAddress As String
Dim strCENTRALFTPUsername As String
Dim strCENTRALFTPPassword As String
Dim strCENTRALFTPFolder As String
'Dim strClientCode As String
Dim lngDelayInSeconds As Long
Private strUsername As String
Private strPassword As String
Private oConfig As a_Nielsen
Private strApplicationDatabase As String
Private flgDbObjConnected As Boolean
Dim strComputerName As String
Private cnWR As ADODB.Connection
Private cnPapyShort As ADODB.Connection
Dim mCL As String
Dim bBOOKFINDISBN13ENABLED As Boolean
Dim bInternetDialup As Boolean

Private mCurrentSecurityCode As String
Event RecordAdded(Success As Boolean)
Event STATUS(pMsg As String)
Dim tlEntities As z_TextListSimple
Dim strDefaultAreaCode As String
Dim strBookfindRoot As String
Private oBookfindmanager As z_BookfindManager
Private bBFLoaded As Boolean
Private lngUnallocatedPT As Long
Private dblVATRATE As Double
Dim strBookfindFacet As String
Dim rsProperty As New ADODB.Recordset


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Sub setDBname(Val As String)
    On Error GoTo errHandler
    mCL = Val
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.setDBname(val)", Val, EA_NORERAISE
    handleError
End Sub

Public Property Get BF() As z_BookfindManager
    Set BF = oBookfindmanager
End Property
Public Property Get BookFindFacet() As String
    BookFindFacet = strBookfindFacet
End Property

Public Property Get UnallocatedPT() As Long
    UnallocatedPT = lngUnallocatedPT
End Property
Public Property Get VATRate() As Long
    VATRate = dblVATRATE
End Property

Public Property Get Configuration() As a_Nielsen
    Set Configuration = oConfig
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
    
    dteTimeStarted = Now()
    Set oPC = Me
    
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.Class_Initialize", , EA_NORERAISE
    handleError
End Sub
Public Property Get CO() As ADODB.Connection
    Set CO = cnWR
End Property
Public Property Get COShort() As ADODB.Connection
    On Error GoTo errHandler
    Set COShort = cnPapyShort
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.COShort", , EA_NORERAISE
    handleError
End Property
Public Property Get Entities() As z_TextListSimple
    Set Entities = tlEntities
End Property
Public Function LoadInitialData(pCommandLine As String) As Boolean
    On Error GoTo errHandler
Dim iConnect As Integer
Dim i As Integer
Dim rs As ADODB.Recordset

    RaiseEvent STATUS("Loading configuration data . . .")
    If iConnect = 0 Then 'Connected to DB
        Set oConfig = New a_Nielsen
        oConfig.Load
    End If
    Set rs = New ADODB.Recordset
    rs.Open "SELECT dbo.tStore.STORE_Code FROM dbo.tConfiguration INNER JOIN dbo.tStore ON dbo.tConfiguration.CF_DefaultStoreID = dbo.tStore.STORE_ID", cnPapyShort, adOpenStatic
    If Not rs.EOF And Not rs.BOF Then
        strClientCode = FNS(rs.Fields(0))
    End If
    rs.Close
    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT CF_UNALLOCATEDPT,CF_VATRATE FROM tConfiguration ", cnPapyShort, adOpenStatic
    If Not rs.EOF And Not rs.BOF Then
        lngUnallocatedPT = FNN(rs.Fields(0))
        dblVATRATE = FNDBL(rs.Fields(1))
    End If
    rs.Close
    Set rs = Nothing
    
    If UCase(pCommandLine) = "FETCH" Then  'only need to connect to bookfind when fetchinf stock sharing data
        If strBookfindRoot > "" Then
            Set oBookfindmanager = New z_BookfindManager
            If Not oBookfindmanager.PrepareBookfind() Then
                iConnect = 2 'Connected to DB but not to Bookfind
            End If
            bBFLoaded = oBookfindmanager.BookfindConnected
        End If
    End If
    If iConnect = 0 Or iConnect = 2 Then flgDbObjConnected = True
    RaiseEvent STATUS("")
    
EXIT_Handler:
    LoadInitialData = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.LoadInitialData(pCommandLine)", pCommandLine, EA_NORERAISE
    handleError
End Function
Public Function Connection() As ADODB.Connection
    On Error GoTo errHandler
    Set Connection = cnWR
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.Connection", , EA_NORERAISE
    handleError
End Function
'Private Function OpenDB() As Integer
'    On Error GoTo errHandler
'Dim strConn As String
'    strPassword = ""
'    strUsername = "sa"
'    OpenDB = 0
'    If cnWR Is Nothing Then
'        Set cnWR = New ADODB.Connection
'        cnWR.Provider = "sqloledb"
'        cnWR.ConnectionTimeout = 10
'        cnWR.Open "Data Source=" & strServerName & ";Initial Catalog=PBKS;User Id=" & strUsername & ";Password=" & strPassword & "; Network Library=dbmssocn;"
'    End If
'    Exit Function
'
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_connection.OpenDB"
'End Function

Public Property Get Connected() As Boolean
    Connected = flgDbObjConnected
End Property
Public Property Let Username(Val As String)
    strUsername = Val
End Property
Public Property Get Username() As String
    Username = strUsername
End Property
Public Property Let Password(Val As String)
    strPassword = Val
End Property
Public Property Get Password() As String
    Password = strPassword
End Property

Public Property Get TimeStarted() As Date
    On Error GoTo errHandler
    TimeStarted = Format(dteTimeStarted, "dd/mm/yyyy hh:mm")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.TimeStarted", , EA_NORERAISE
    handleError
End Property
Public Sub Disconnect()
    On Error GoTo errHandler
    Set oConfig = Nothing
    If flgDbObjConnected Then
        cnWR.Close
        Set cnWR = Nothing
        flgDbObjConnected = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.Disconnect", , EA_NORERAISE
    handleError
End Sub

Public Property Let CurrentSecurityCode(Val As String)
    mCurrentSecurityCode = Val
End Property
Public Property Get CurrentSecurityCode() As String
    CurrentSecurityCode = mCurrentSecurityCode
End Property

Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set tlEntities = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.Class_Terminate", , EA_NORERAISE
    handleError
End Sub
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
    ErrorIn "a_connection.NameOfPC", , EA_NORERAISE
    handleError
End Property


Public Sub InitializeSettings()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String
Dim strPCName As String

    
    strPCName = Trim(Me.NameOfPC)
    
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
    End If
    
    strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
    strServerName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
    strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")
    
    LoadProperties
    
    strAdministratorEmail = GetProperty("ADMINISTRATOREMAIL")
    strSMTPServer = GetProperty("SMTPServer")
    strSMTPUsername = GetProperty("SMTP_Username")
    strSMTPPassword = GetProperty("SMTP_Password")
    strCENTRALFTPAddress = GetProperty("CENTRALFTPADDRESS")
    strCENTRALFTPUsername = GetProperty("CENTRALFTPUSERNAME")
    strCENTRALFTPPassword = GetProperty("CENTRALFTPPASSWORD")
    strCENTRALFTPFolder = GetProperty("CENTRALFTPFOLDER")
    lngDelayInSeconds = GetProperty("DELAYINSECONDS")
    strDefaultAreaCode = GetProperty("DEFAULTAREACODE")
    strBookfindRoot = GetProperty("BOOKFINDROOT")
    strBookfindFacet = GetProperty("BOOKFINDFACET")
    bBOOKFINDISBN13ENABLED = GetProperty("BOOKFINDISBN13ENABLED")
    bInternetDialup = GetProperty("INTERNETDIALUP") = "TRUE"         '(GetIniKeyValue(strServerMachineSharedFolder & "\PBKS.INI", "SUPPORT", "INTERNETDIALUP", "TRUE") = "TRUE")
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.InitializeSettings", , EA_NORERAISE
    handleError
End Sub
Public Function OpenDBSHort(Optional bRestart As Boolean) As Integer
    On Error GoTo errHandler
Dim strMainConnectionString As String

    OpenDBSHort = 0
    If cnPapyShort Is Nothing Then
        Set cnPapyShort = New ADODB.Connection
        cnPapyShort.Provider = "sqloledb"
        cnPapyShort.CommandTimeout = 60
        strMainConnectionString = "Data Source=" & Me.servername & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & "; Connect Timeout=45"
        If mCL = "" Then
            cnPapyShort.Open strMainConnectionString
        Else
            cnPapyShort.Open "Data Source=" & Me.servername & ";Initial Catalog=" & mCL & ";User Id=sa;Password=" & strPassword & "; Connect Timeout=45"
        End If
        cnPapyShort.CommandTimeout = 0
        OpenDBSHort = 0
    Else
        If cnPapyShort.Errors.Count > 0 Then
            
            cnPapyShort.Close
            If mCL = "" Then
                cnPapyShort.Open strMainConnectionString
            Else
                cnPapyShort.Open "Data Source=" & Me.servername & ";Initial Catalog=" & mCL & ";User Id=sa;Password=" & strPassword & ";Connect Timeout=45"
            End If
        Else
            OpenDBSHort = 99
        End If
    End If

EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.OpenDBSHort(bRestart)", bRestart, EA_NORERAISE
End Function
Public Sub DisconnectDBShort()
    On Error GoTo errHandler
    cnPapyShort.Close
    Set cnPapyShort = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.DisconnectDBShort", , EA_NORERAISE
End Sub

Public Property Get BookFindRoot() As String
    BookFindRoot = strBookfindRoot
End Property

Public Property Get servername() As String
    servername = strServerName
End Property
Public Property Get ServerMachineName() As String
    ServerMachineName = strPBKSSERVERMACHINE
End Property

Public Property Get SMTPServer() As String
    SMTPServer = strSMTPServer
End Property
Public Property Get AdministratorEmail() As String
    AdministratorEmail = strAdministratorEmail
End Property
Public Property Get ServerRootFolder() As String
    On Error GoTo errHandler
    ServerRootFolder = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.ServerRootFolder", , EA_NORERAISE
End Property
'Public Property Get ClientName() As String
'    ClientName = FNS(strClientName)
'End Property

Public Property Get ClientCode() As String
    On Error GoTo errHandler
    ClientCode = FNS(strClientCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.ClientCode", , EA_NORERAISE
End Property

Public Property Get SharedFolderRoot() As String
    SharedFolderRoot = strSharedFolderRoot
End Property

Public Property Get CentralFTPAddress() As String
    On Error GoTo errHandler
    CentralFTPAddress = FNS(strCENTRALFTPAddress)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.CentralFTPAddress", , EA_NORERAISE
End Property
Public Property Get CentralFTPUsername() As String
    On Error GoTo errHandler
    CentralFTPUsername = FNS(strCENTRALFTPUsername)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.CentralFTPUsername", , EA_NORERAISE
End Property
Public Property Get CentralFTPPassword() As String
    On Error GoTo errHandler
    CentralFTPPassword = FNS(strCENTRALFTPPassword)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.CentralFTPPassword", , EA_NORERAISE
End Property
Public Property Get CentralFTPFolder() As String
    On Error GoTo errHandler
    CentralFTPFolder = FNS(strCENTRALFTPFolder)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.CentralFTPFolder", , EA_NORERAISE
End Property
Public Property Get CentralFTPPassive() As Boolean
    CentralFTPPassive = True
End Property
Public Property Get DelayInSeconds() As Long
    DelayInSeconds = lngDelayInSeconds
End Property
Public Property Get DefaultAreaCode() As String
    DefaultAreaCode = strDefaultAreaCode
End Property
Public Property Get BOOKFINDISBN13ENABLED() As Boolean
    BOOKFINDISBN13ENABLED = bBOOKFINDISBN13ENABLED
End Property
Public Property Get InternetDialup() As Boolean
    InternetDialup = bInternetDialup
End Property

Public Function GetProperty(pKey As String) As String
    On Error GoTo errHandler
    rsProperty.MoveFirst
    rsProperty.Find "PropertyKey = '" & pKey & "'"
    GetProperty = Trim(CStr(rsProperty.Fields(1)))
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.GetProperty(pKey)", pKey, EA_NORERAISE
End Function
Public Function LoadProperties() As Boolean
    On Error GoTo errHandler
Dim sSQL As String
    
 Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
   
    sSQL = "SELECT * FROM tProperty"
    Set rsProperty = New ADODB.Recordset
    rsProperty.CursorLocation = adUseClient
    rsProperty.Open sSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    Set rsProperty.ActiveConnection = Nothing
'-------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.LoadProperties", , EA_NORERAISE
End Function

Public Function DatabaseName() As String
    On Error GoTo errHandler
    DatabaseName = "PBKS"
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.DatabaseName", , EA_NORERAISE
End Function
Public Function WorkstationID() As Long
    On Error GoTo errHandler
    WorkstationID = 0
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_connection.WorkstationID", , EA_NORERAISE
End Function

