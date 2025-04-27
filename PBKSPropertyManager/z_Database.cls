VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_POSConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DBConn As ADODB.Connection
Dim strPassword As String
Dim strUsername As String
Dim strServerRootPath As String
Dim strServer As String
Dim strSharedFolderRoot As String

Dim strBackupFolder As String
Dim strDDB As String
Dim sLocalMachineName As String
Dim strLocalRootFolder As String
Dim strMainSQLServerName As String
Dim strMainConnectionString As String
Dim strDBName As String
Dim strPBKSSERVERMACHINE As String
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Property Get ConnectionString() As String
    ConnectionString = strMainConnectionString
End Property
Public Property Let DatabaseName(val As String)
    strDBName = val
End Property
Public Function dbConnect(pstrServer As String, pPassword As String) As Integer
    On Error GoTo errHandler

    
    If strDBName = "" Then
        strDBName = "PBKS"
    End If
    
    If DBConn Is Nothing Then
        Set DBConn = New ADODB.Connection
    End If
    If DBConn.State <> adStateClosed Then
        DBConn.Close
    End If
        If pstrServer = "PBKS-SVR" Then
            strUsername = "Psa"
            strPassword = "jadarow"
        Else
            strUsername = "sa"
            strPassword = "car"
        End If

   ' strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=" & pstrServer & ";Initial Catalog=" & strDBName & ";User Id=sa;Password=" & strPassword & "; Connect Timeout=45"
    strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & pstrServer & ";Initial Catalog=" & strDBName & ";User Id= " & strUsername & ";Password=" & strPassword & ";Connect Timeout=50"
    DBConn.Open strMainConnectionString
    If DBConn.State = adStateClosed Then
        dbConnect = 99
    Else
        dbConnect = 0
    End If
    
    
    
    
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSConnection.dbConnect(pstrServer)", pstrServer, , , "strMainConnectionString", Array(strMainConnectionString)
End Function
Public Function dbCloseConnect() As Integer
    On Error Resume Next

    DBConn.Close
    Set DBConn = Nothing
    
EXIT_Handler:
    Exit Function

    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSConnection.dbCloseConnect"
End Function
Private Function NameOfPC() As String
    On Error GoTo errHandler
Dim NameSize As Long
Dim MachineName As String * 16
Dim x As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    x = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSConnection.NameOfPC"
End Function

Private Sub InitializeSettings()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String

    sLocalMachineName = NameOfPC
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & sLocalMachineName & "\PBKS_S"
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
        strSharedFolderRoot = "C:\PBKS"
    End If
    
    strMainSQLServerName = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSConnection.InitializeSettings"
End Sub
Public Property Get LocalRoot() As String

    LocalRoot = strLocalRootFolder
End Property
Public Property Get SharedFolderRoot()
    SharedFolderRoot = strSharedFolderRoot
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
    InitializeSettings
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSConnection.Class_Initialize"
End Sub

