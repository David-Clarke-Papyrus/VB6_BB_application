Attribute VB_Name = "mConnectionWizard"
Option Explicit

Global oCnn As ADODB.Connection
Global strSharedFolderRoot As String
Global strLocalRootFolder As String
Global bUseTest As Boolean
Global strPBKSSERVERMACHINE As String
Global strServername As String
Global strDatabaseName As String
Global strPassword As String
Global strMainConnectionString As String
Global strPCName As String

Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Public Sub InitializeSettings(Optional bUseTest As Boolean)
Dim fs As New FileSystemObject
Dim strTag As String
Dim strTmp As String
Dim strValue As String
Dim strRootPath  As String

    strPCName = Trim(NameOfPC)
    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
        If bUseTest Then
            strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSERVERMACHINE", strPCName)
        Else
            strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        End If
        strSharedFolderRoot = "\\" & strPBKSSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        If bUseTest Then
            strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSERVERMACHINE", strPCName)
        Else
            strPBKSSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PBKSSERVERMACHINE", strPCName)
        End If
        strSharedFolderRoot = "C:\PBKS"
    End If
    If bUseTest Then
        strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "TESTSQLSERVER", "")
        strDatabaseName = "PBKSTEST"
    Else
        strServername = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "MAINSQLSERVER", "")
        strDatabaseName = "PBKS"
    End If
    strPassword = GetIniKeyValue(strLocalRootFolder & "\PBKSWS.INI", "NETWORK", "PASSWORD", "")
    
    
End Sub

Public Property Get NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Property
End Property


Sub ConnectToDatabase()
    InitializeSettings
    Set oCnn = New ADODB.Connection
    oCnn.Provider = "sqloledb"
    oCnn.CommandTimeout = 60
    strMainConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=" & strServername & ";Initial Catalog=" & strDatabaseName & ";User Id=sa;Password=" & strPassword & ";Connect Timeout=45"

    oCnn.Open strMainConnectionString
    oCnn.CommandTimeout = 0

End Sub
