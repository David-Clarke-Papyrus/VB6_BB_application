VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_CentralConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim strMainConnectionString As String
Private strUsername As String
Private bBFLoaded As Boolean
Private oConfig As a_c_Configuration
Private oBookfindmanager As z_BookfindManager
Private strPassword As String
Private cnPapyShort As ADODB.Connection
Private strServerName As String
Private strCENTRALSERVERMACHINE As String
Private strRemovableBackupShareName As String
Private strCompressBackup As String
Private strBackupMedium As String
Dim strSharedFolderRoot As String
Private strDDB As String
Private strBookfindRoot As String
Private strLocalRootFolder As String
Private strCENTRALFTPAddress As String
Private strCENTRALFTPFolder As String
Private strCENTRALFTPPassword As String
Private strCENTRALFTPUsername As String
Private strDefaultAreaCode As String
Dim bBOOKFINDISBN13ENABLED As Boolean

Private strDUN As String
Private strBookfindFacet As String
Dim mlngWSID As Long
Dim strPCName As String
Dim bEDIEnabled As Boolean
Dim mCL As String
Dim mINI As String


Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Property Get ConnectionString() As String
    ConnectionString = strMainConnectionString
End Property
Public Property Get BOOKFINDISBN13ENABLED() As Boolean
    BOOKFINDISBN13ENABLED = bBOOKFINDISBN13ENABLED
End Property

Public Property Get BF() As z_BookfindManager
    Set BF = oBookfindmanager
End Property
Public Property Get Configuration() As a_c_Configuration
    Set Configuration = oConfig
End Property
'Private Sub objBookfindManager_RecordAdded(Success As Boolean)
'    RaiseEvent RecordAdded(Success)
'End Sub
Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    'InitializeSettings mINI
    Set oPC = Me
EXIT_Handler:
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.Class_Initialize"
End Sub
'Public Property Get CO() As ADODB.Connection
'    Set CO = cnCentral
'End Property
Public Sub DisconnectFromBookfind()
On Error Resume Next
    If Not oBookfindmanager Is Nothing Then
        oBookfindmanager.CloseDB
        Set oBookfindmanager = Nothing
    End If
End Sub

Public Function LoadInitialData(INI As String) As Boolean
    On Error GoTo ErrHandler
Dim iConnect As Integer
Dim i As Integer
Dim strPos As String
Dim OpenResult As Integer

    mINI = INI
    LoadInitialData = True
    Set oPC = Me
    Set cnPapyShort = Nothing
    bBFLoaded = False

'-------------------------------
    OpenResult = OpenDBSHort
'-------------------------------
        Set oConfig = New a_c_Configuration
        oConfig.Load
'---------------------------------------------------
    If OpenResult = 0 Then DisconnectDBShort
'---------------------------------------------------
        i = 1
        Do While oPC.Configuration.LookupSequence(i) > ""
            Select Case oPC.Configuration.LookupSequence(i)
            Case "BF"
                    Set oBookfindmanager = New z_BookfindManager
                    If Not oBookfindmanager.PrepareBookfind() Then
                        iConnect = 2 'Connected to DB but not to Bookfind
                    Else
                        bBFLoaded = True
                    End If
            Case "WH"
            End Select
            i = i + 1
        Loop
    
    
EXIT_Handler:
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.Connect()", , , , strPos, Array(strPos)
End Function

Public Property Get COShort() As ADODB.Connection
    On Error GoTo ErrHandler
    Set COShort = cnPapyShort
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.COShort"
End Property
Public Sub SetConnectionString(pCL As String)
    strMainConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=sa;Data Source=" & Me.servername & ";Initial Catalog=PBKSC;User Id=sa;Password=" & strPassword & ";Connect Timeout=120"
    If pCL > "" Then
        strMainConnectionString = Replace(strMainConnectionString, "PBKSC", pCL)
    End If

End Sub
Public Function OpenDBSHort(Optional restart As Boolean) As Integer
Dim strPos As String
Dim iCnt As Integer
Dim tmprs As ADODB.Recordset

    iCnt = 0
    On Error GoTo ErrHandler
    
    OpenDBSHort = 0
    If cnPapyShort Is Nothing Then
        Set cnPapyShort = New ADODB.Connection
    End If
    If cnPapyShort Is Nothing Then
        Set cnPapyShort = New ADODB.Connection
    End If
    
    If mCL = "" Then
        mCL = "PBKSC"
    End If
TryAgain:
    If cnPapyShort.State = adStateClosed Or restart Then
        strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & Me.servername & ";Initial Catalog=" & mCL & ";User Id=sa;Password=" & strPassword & ";Connect Timeout=40"
        cnPapyShort.Open strMainConnectionString
        cnPapyShort.CommandTimeout = 240
        
    Else
        Set tmprs = cnPapyShort.execute("SELECT CF_VATRATE FROM tConfiguration")
        tmprs.Close
        OpenDBSHort = 99
    End If
    

EXIT_Handler:
    Exit Function
ErrHandler:
    ErrPreserve
    If Err = -2147467259 And iCnt < 2 Then
        If cnPapyShort.State <> adStateClosed Then cnPapyShort.Close
        cnPapyShort.errors.Clear
        iCnt = iCnt + 1
        GoTo TryAgain
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.OpenDBSHort", , , , "strMainConnectionString:strPOS", Array(strMainConnectionString, strPos)
End Function
Public Sub DisconnectDBShort()
    On Error GoTo ErrHandler
    cnPapyShort.Close
    Set cnPapyShort = Nothing
    Exit Sub
ErrHandler:
    ErrPreserve
    If Err = 91 Then 'Already closed
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.DisconnectDBShort"
End Sub

Public Property Get BFLoaded() As Boolean
    BFLoaded = bBFLoaded
End Property
Public Property Let Username(val As String)
    strUsername = val
End Property
Public Property Get Username() As String
    Username = strUsername
End Property
Public Property Get CentralServerComputername() As String
   CentralServerComputername = strCENTRALSERVERMACHINE
End Property
Public Sub InitializeSettings(INI As String)
Dim fs As New FileSystemObject
Dim strRootPath  As String

On Error GoTo ERRH

Dim strPCName As String
    
    strPCName = Trim(Me.NameOfPC)
    mINI = INI
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
        strCENTRALSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\" & INI, "NETWORK", "CENTRALSERVERMACHINE", "")
        strSharedFolderRoot = "\\" & strCENTRALSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strSharedFolderRoot = strLocalRootFolder
    End If
    
    strPassword = GetIniKeyValue(strLocalRootFolder & "\" & INI, "NETWORK", "PASSWORD", "")
    strServerName = GetIniKeyValue(strLocalRootFolder & "\" & INI, "NETWORK", "MAINSQLSERVER", "")
    If strServerName = "" Then
        strServerName = "MARIO\PBKSINSTANCE"
    End If
    strRemovableBackupShareName = GetIniKeyValue(strLocalRootFolder & "\" & INI, "FOLDERS", "BACKUP", "")
    
    strCompressBackup = GetIniKeyValue(strLocalRootFolder & "\" & INI, "OPTIONS", "BACKUPCOMPRESSION ", "")
    strBackupMedium = GetIniKeyValue(strLocalRootFolder & "\" & INI, "OPTIONS", "BACKUPMEDIUM ", "")
    
    strDDB = GetIniKeyValue(strLocalRootFolder & "\" & INI, "BOOKDATA", "BFDISTRIBUTORDB", "")
    strBookfindRoot = GetIniKeyValue(strLocalRootFolder & "\" & INI, "BOOKDATA", "BOOKFINDROOT", "")
    strBookfindFacet = GetIniKeyValue(strLocalRootFolder & "\" & INI, "BOOKDATA", "BOOKFINDFACET", "WEBK")
    strCENTRALFTPAddress = GetIniKeyValue(strLocalRootFolder & "\" & INI, "CENTRAL", "FTPADDRESS", "")
    strCENTRALFTPFolder = GetIniKeyValue(strLocalRootFolder & "\" & INI, "CENTRAL", "FTPFolder", "")
    strCENTRALFTPUsername = GetIniKeyValue(strLocalRootFolder & "\" & INI, "CENTRAL", "FTPUSERNAME", "")
    strCENTRALFTPPassword = GetIniKeyValue(strLocalRootFolder & "\" & INI, "CENTRAL", "FTPPassword", "")
    strDUN = GetIniKeyValue(strLocalRootFolder & "\" & INI, "CENTRAL", "CONNECTIONNAME", "")
    strDefaultAreaCode = GetIniKeyValue(strLocalRootFolder & "\" & INI, "LOCAL", "DEFAULTAREACODE", "  ")
    bBOOKFINDISBN13ENABLED = GetIniKeyValue(strLocalRootFolder & "\" & INI, "BOOKDATA", "BOOKFINDISBN13ENABLED", "FALSE") = "TRUE"
ERRH:
    Exit Sub
    'Resume
End Sub
Public Property Get servername() As String
    servername = strServerName
End Property

Public Property Get NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
End Property
'
Public Property Let Password(val As String)
    strPassword = val
End Property

Public Property Get BookFindFacet() As String
    BookFindFacet = strBookfindFacet
End Property
Public Property Get BookFindRoot() As String
    BookFindRoot = strBookfindRoot
End Property
Public Property Get RemovableBackupShareName() As String
    RemovableBackupShareName = "\\" & NameOfPC & "\" & strRemovableBackupShareName & "\"
End Property
Public Property Get RemoveableBackupMedium() As String
    RemoveableBackupMedium = strBackupMedium
End Property
Public Property Get CompressBackup() As String
    CompressBackup = strCompressBackup
End Property
Public Property Get CentralFTPAddress() As String
    CentralFTPAddress = strCENTRALFTPAddress
End Property
Public Property Get CentralFTPFolder() As String
    CentralFTPFolder = strCENTRALFTPFolder
End Property
Public Property Get CentralFTPUsername() As String
    CentralFTPUsername = strCENTRALFTPUsername
End Property
Public Property Get CentralFTPPassword() As String
    CentralFTPPassword = strCENTRALFTPPassword
End Property
Public Property Get DUN() As String
    DUN = strDUN
End Property

Public Property Get SharedFolderRoot() As String
    SharedFolderRoot = strSharedFolderRoot
End Property
Public Property Get DistributorDB() As String
    DistributorDB = strDDB
End Property
Public Property Get BackupFolder() As String
    BackupFolder = strLocalRootFolder & "\BU\"
End Property

Private Sub NotifyConnection()
    On Error GoTo ErrHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_NotifyConnection"
    cmd.CommandType = adCmdStoredProc
    
    ' Get parameter value and append parameter.
    Set prm = cmd.CreateParameter("@COMPNAME", adVarChar, adParamInput, 30, Me.NameOfPC)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("@WSID", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    cmd.execute
    
    
    mlngWSID = cmd.Parameters("@WSID").Value
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.NotifyCOnnection"
End Sub
Public Property Get WorkstationID() As Long
    WorkstationID = mlngWSID
End Property
Public Property Get WorkstationName() As String
    WorkstationName = strPCName
End Property
Public Property Get DefaultAreaCode() As String
    DefaultAreaCode = strDefaultAreaCode
End Property

Public Property Get CatalogName() As String
    If mCL = "" Then
        CatalogName = "PBKSC"
    Else
        CatalogName = mCL
    End If

End Property
Public Property Get Password() As String
    Password = strPassword
End Property
Public Property Get SupportsCatalogue() As Boolean
    SupportsCatalogue = False
End Property

