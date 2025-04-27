VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PapyConnRep"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private dteTimeStarted As Date
Private strPapyConnErr As String
Private strUsername As String
Private strPWD As String
Private strDatabase As String
Private oConfig As a_Configuration
Private strApplicationDatabase As String
Private bBFLoaded As Boolean
Private flgDbObjConnected As Boolean
Private strDatabaseFullPath As String
Private cnPapy As ADODB.Connection
Event RecordAdded(Success As Boolean)
Event Status(pMsg As String)
Private strServername As String
Dim strLocalRootFolder As String
Dim strSharedFolderRoot As String
Dim strRemovableBackupShareName As String
Dim strPBKSSERVERMACHINE As String
Dim strCompressBackup As String
Dim strBackupMedium As String
Dim mlngWSID As Long
Dim strDefaultAreaCode As String
Dim strInternetDialup As String
Dim strConnectionName As String

Private mCurrentSecurityCode As String

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Property Get Configuration() As a_Configuration
    Set Configuration = oConfig
End Property
Private Sub objBookfind_RecordAdded(Success As Boolean)
    On Error GoTo errHandler
    RaiseEvent RecordAdded(Success)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.objBookfind_RecordAdded(Success)", Success
End Sub
Private Sub Class_Initialize()
    On Error GoTo errHandler
    dteTimeStarted = Now()
    Set oPC = Me
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.Class_Initialize"
End Sub
Public Property Get CO() As ADODB.Connection
    Set CO = cnPapy
End Property

Public Function Connect(Optional pDBName As String) As Long
    On Error GoTo errHandler
Dim iConnect As Integer
Dim i As Integer

    Set oPC = Me
    If Not flgDbObjConnected Then
        Set cnPapy = Nothing
        RaiseEvent Status("Opening database connection . . .")
        iConnect = OpenDB(pDBName)
    End If
    RaiseEvent Status("Loading configuration data . . .")
    If iConnect = 0 Then 'Connected to DB
        Set oConfig = New a_Configuration
        oConfig.Load
    End If
    
    If iConnect = 0 Or iConnect = 2 Then flgDbObjConnected = True
    RaiseEvent Status("")
    
EXIT_Handler:
    Connect = iConnect
    Exit Function
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.Connect(pDBName)", pDBName
End Function

Private Function OpenDB(pCL As String) As Integer
    On Error GoTo errHandler
Dim strMainConnectionString As String

    OpenDB = 0
    If cnPapy Is Nothing Then
        Set cnPapy = New ADODB.Connection
'        Set cnPapyShort = New ADODB.Connection
        cnPapy.Provider = "sqloledb"
        cnPapy.CommandTimeout = 60
            strMainConnectionString = "Data Source=" & strServername & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPWD & "; Network Library=dbmssocn;Connect Timeout=45"
        If pCL = "" Then
            cnPapy.Open strMainConnectionString
        Else
            cnPapy.Open "Data Source=" & strServername & ";Initial Catalog=" & pCL & ";User Id=" & strUsername & ";Password=" & strPWD & "; Network Library=dbmssocn;Connect Timeout=45"
        End If
        cnPapy.CommandTimeout = 0
    Else
    End If
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.OpenDB(pCL)", pCL
End Function

Public Property Get Connected() As Boolean
    Connected = flgDbObjConnected
End Property
Public Property Get BFLoaded() As Boolean
    BFLoaded = bBFLoaded
End Property
Public Property Let UserName(val As String)
    strUsername = val
End Property
Public Property Get UserName() As String
    UserName = strUsername
End Property
Public Property Let Password(val As String)
    strPWD = val
End Property
Public Property Let Database(val As String)
    strDatabase = val
End Property
Public Property Let ApplicationDatabase(val As String)
    strApplicationDatabase = val
End Property

Public Property Get TimeStarted() As Date
    On Error GoTo errHandler
    TimeStarted = Format(dteTimeStarted, "dd/mm/yyyy hh:mm")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.TimeStarted"
End Property
Public Sub Disconnect()
    On Error GoTo errHandler
    Set oConfig = Nothing
    If flgDbObjConnected Then
        cnPapy.Close
        Set cnPapy = Nothing
        flgDbObjConnected = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.Disconnect"
End Sub

Public Sub InitializeSettings()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strPCName As String

    strPCName = Trim(Me.NameOfPC)
    
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        strSharedFolderRoot = "C:\PBKS"
    End If
    
    
    strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
    strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", strPCName)
    
    strRemovableBackupShareName = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "FOLDERS", "BACKUP", "")
    strCompressBackup = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "BACKUPCOMPRESSION ", "TRUE")
    strBackupMedium = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "OPTIONS", "BACKUPMEDIUM ", "")
    strDefaultAreaCode = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "LOCAL", "DEFAULTAREACODE ", "")
    strInternetDialup = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "SUPPORT", "INTERNETDIALUP ", "YES")
    strConnectionName = GetIniKeyValue(strSharedFolderRoot & "\PBKS.INI", "SUPPORT", "CONNECTIONNAME ", "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.InitializeSettings"
End Sub
Public Property Get InternetDialup() As Boolean
    InternetDialup = (strInternetDialup = "YES")
End Property
Public Property Get Connectionname() As String
    Connectionname = strConnectionName
End Property
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
    ErrorIn "PapyConn.NameOfPC"
End Property
Public Property Get SharedFolderRoot() As String
    On Error GoTo errHandler
    SharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.SharedFolderRoot"
End Property
Public Property Get WorkstationID() As Long
    WorkstationID = mlngWSID
End Property
Public Property Let CurrentSecurityCode(val As String)
    mCurrentSecurityCode = val
End Property
Public Property Get CurrentSecurityCode() As String
    CurrentSecurityCode = mCurrentSecurityCode
End Property
Public Property Get DefaultAreaCode() As String
    DefaultAreaCode = strDefaultAreaCode
End Property

