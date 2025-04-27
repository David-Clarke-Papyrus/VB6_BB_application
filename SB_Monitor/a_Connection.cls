VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strMainConnectionString As String
Private strUsername As String
Private strBranchCode As String
Private strPassword As String
Private cnn As ADODB.Connection
Private strServerName As String
Private strServerMachine As String
Dim strLocalRootFolder As String

Dim strSharedFolderRoot As String
Private strDDB As String



Dim mlngWSID As Long
Dim strPCName As String

Dim mCL As String


Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Sub SetBranchCode(val As String)
    strBranchCode = Trim(val)
End Sub

Public Property Get ConnectionString() As String
    ConnectionString = strMainConnectionString
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
'    InitializeSettings
'    SetConnectionString ""
    Set oCnn = Me
EXIT_HANDLER:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Connection.Class_Initialize"
End Sub


Public Property Get Connection() As ADODB.Connection
    On Error GoTo errHandler
    Set Connection = cnn
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.Connection"
End Property
Public Sub SetConnectionString()
    On Error GoTo errHandler
    strMainConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=sa;Data Source=" & Me.ServerName & ";Initial Catalog=PBKS;User Id=sa;Password=" & strPassword & ";Connect Timeout=120"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Connection.SetConnectionString"
End Sub
Public Function OpenDB() As Integer
    On Error GoTo errHandler
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
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.OpenDB"
End Function
Public Sub CloseDB()
    On Error GoTo errHandler
    cnn.Close
    Set cnn = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_HubConnection.CloseDB"
End Sub

Public Property Let Username(val As String)
    strUsername = val
End Property
Public Property Get Username() As String
    Username = strUsername
End Property
Public Sub InitializeSettings()
    On Error GoTo errHandler
Dim strRootPath  As String


Dim strPCName As String
    
    strPCName = Trim(Me.NameOfPC)
    
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
        strServerMachine = GetIniKeyValue(strLocalRootFolder & "\SBMONITOR.INI", "NETWORK", strBranchCode & "MACHINE", "")
        strSharedFolderRoot = "\\" & strServerMachine & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strSharedFolderRoot = strLocalRootFolder
    End If
    
    strPassword = GetIniKeyValue(strLocalRootFolder & "\SBMONITOR.INI", "NETWORK", "PASSWORD", "")
    strServerName = GetIniKeyValue(strLocalRootFolder & "\SBMONITOR.INI", "NETWORK", strBranchCode, "")
  '  strDDB = GetIniKeyValue(strLocalRootFolder & "\SBMONITOR.INI", "BOOKDATA", "BFDISTRIBUTORDB", "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Connection.InitializeSettings"
End Sub
Public Property Get ServerName() As String
    ServerName = strServerName
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
    ErrorIn "a_HubConnection.NameOfPC"
End Property

Public Property Let Password(val As String)
    strPassword = val
End Property
Public Property Get SharedFolderRoot() As String
    SharedFolderRoot = strSharedFolderRoot
End Property
Public Property Get LocalFolderRoot() As String
    LocalFolderRoot = strLocalRootFolder
End Property

Public Property Get Password() As String
    Password = strPassword
End Property


