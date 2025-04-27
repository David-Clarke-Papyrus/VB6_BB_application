VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_HubConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strMainConnectionString As String
Private strUsername As String
Private bBFLoaded As Boolean
Private strPassword As String
Private cnn As ADODB.Connection
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
Private Sub Class_Initialize()
    On Error GoTo ErrHandler
    InitializeSettings
    SetConnectionString ""
    Set oCnn = Me
EXIT_HANDLER:
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.Class_Initialize"
End Sub
Public Sub DisconnectFromBookfind()
    On Error GoTo ErrHandler
    If Not oBookfindmanager Is Nothing Then
        oBookfindmanager.CloseDB
        Set oBookfindmanager = Nothing
    End If
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.DisconnectFromBookfind"
End Sub


Public Property Get Connection() As ADODB.Connection
    On Error GoTo ErrHandler
    Set Connection = cnn
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.Connection"
End Property
Public Sub SetConnectionString(pCL As String)
    On Error GoTo ErrHandler
    strMainConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=sa;Data Source=" & Me.ServerName & ";Initial Catalog=HUB;User Id=sa;Password=" & strPassword & ";Connect Timeout=120"
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.SetConnectionString(pCL)", pCL
End Sub
Public Function OpenDB() As Integer
    On Error GoTo ErrHandler
Dim strPos As String
    
    OpenDB = 0
    If cnn Is Nothing Then
        Set cnn = New ADODB.Connection
    End If
    If cnn.Errors.Count > 0 Or cnn.State = adStateClosed Then
        cnn.Open strMainConnectionString
        cnn.CommandTimeout = 360
        
    Else
        OpenDB = 99
    End If


EXIT_HANDLER:
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.OpenDB"
End Function
Public Sub CloseDB()
    On Error GoTo ErrHandler
    cnn.Close
    Set cnn = Nothing
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.CloseDB"
End Sub

Public Property Get BFLoaded() As Boolean
    BFLoaded = bBFLoaded
End Property
Public Property Let Username(Val As String)
    strUsername = Val
End Property
Public Property Get Username() As String
    Username = strUsername
End Property
Public Sub InitializeSettings()
    On Error GoTo ErrHandler
Dim fs As New FileSystemObject
Dim strRootPath  As String


Dim strPCName As String
    
    strPCName = Trim(Me.NameOfPC)
    
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
        strCENTRALSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\HUB.INI", "NETWORK", "HUBSERVERMACHINE", "")
        strSharedFolderRoot = "\\" & strCENTRALSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strSharedFolderRoot = strLocalRootFolder
    End If
    
    strPassword = GetIniKeyValue(strLocalRootFolder & "\HUB.INI", "NETWORK", "PASSWORD", "")
    strServerName = GetIniKeyValue(strLocalRootFolder & "\HUB.INI", "NETWORK", "MAINSQLSERVER", "")
    strRemovableBackupShareName = GetIniKeyValue(strLocalRootFolder & "\HUB.INI", "FOLDERS", "BACKUP", "")
    
    strDDB = GetIniKeyValue(strLocalRootFolder & "\HUB.INI", "BOOKDATA", "BFDISTRIBUTORDB", "")
    strBookfindRoot = GetIniKeyValue(strLocalRootFolder & "\HUB.INI", "BOOKDATA", "BOOKFINDROOT", "")
    strBookfindFacet = GetIniKeyValue(strLocalRootFolder & "\HUB.INI", "BOOKDATA", "BOOKFINDFACET", "WEBK")
    bBOOKFINDISBN13ENABLED = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "BOOKDATA", "BOOKFINDISBN13ENABLED", "FALSE") = "TRUE"
    Exit Sub
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.InitializeSettings"
End Sub
Public Property Get ServerName() As String
    ServerName = strServerName
End Property

Public Property Get NameOfPC() As String
    On Error GoTo ErrHandler
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.NameOfPC"
End Property
'
Public Property Let Password(Val As String)
    strPassword = Val
End Property

Public Property Get BookFindFacet() As String
    BookFindFacet = strBookfindFacet
End Property
Public Property Get BookFindRoot() As String
    BookFindRoot = strBookfindRoot
End Property
Public Property Get SharedFolderRoot() As String
    SharedFolderRoot = strSharedFolderRoot
End Property
Public Property Get DistributorDB() As String
    DistributorDB = strDDB
End Property
Public Property Get BackupFolder() As String
    On Error GoTo ErrHandler
    BackupFolder = strLocalRootFolder & "\BU\"
    Exit Property
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.BackupFolder"
End Property

Public Property Get Password() As String
    Password = strPassword
End Property

