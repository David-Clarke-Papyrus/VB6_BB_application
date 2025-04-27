Attribute VB_Name = "Default"
Option Explicit

Global cn As ADODB.Connection
Global strSharedServerFolder As String
Global oTF As New z_TextFile
Global oFSO As FileSystemObject
Global strServerName As String
Global gRes
Global strARG As String
Global strLogPath As String
'FTP
Global FTP1 As FTPClass
Global ftpFile As FTPFileClass
'FTP source
Global gFTPSourceAddress As String
Global gFTPSourceUsername As String
Global gFTPSourcePassword As String
Global gFTPSourceFolder As String

Global gFTPTargetAddress As String
Global gFTPTargetUsername As String
Global gFTPTargetPassword As String
Global gFTPTargetFolder As String

Global gDownloadFolder As String
Global gBackupFolder As String
Global gStoreCode As String
Global gDialup As String
Global strPCName As String
Global fINET As wininet
Global gConnectionName As String
Global gPassword As String
Global strConnection As String

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmMain
    
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    strSharedServerFolder = "C:\PBKS"
    strARG = Command()
    strServerName = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "NETWORK", "SQLSERVERNAME", "")
    gFTPSourceAddress = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "FTPADDRESS", "")
    gFTPSourceUsername = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "FTPUSERNAME", "")
    gFTPSourcePassword = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "FTPPASSWORD", "")
    gFTPSourceFolder = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "FTPFOLDER", "")
    gFTPTargetAddress = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "CENTRAL", "FTPADDRESS", "")
    gFTPTargetUsername = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "CENTRAL", "FTPUSERNAME", "")
    gFTPTargetPassword = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "CENTRAL", "FTPPASSWORD", "")
    gFTPTargetFolder = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "CENTRAL", "FTPFOLDER", "")
    gDownloadFolder = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "DOWNLOADFOLDER", "")
    gBackupFolder = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "BACKUPFOLDER", "BU")
    gStoreCode = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "NETWORK", "STORECODE", "")
    gDialup = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "INTERNETDIALUP", "TRUE")
    gConnectionName = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "INTERNETDIALUP", "TRUE")
    strPCName = Trim(NameOfPC)

    gPassword = GetIniKeyValue(strSharedServerFolder & "\PBKS_WSTOCK.INI", "NETWORK", "PASSWORD", "")
    If cn Is Nothing Then
        Set cn = New ADODB.Connection
        cn.Provider = "sqloledb"
        cn.ConnectionTimeout = 60
        strConnection = "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & strServerName & ";Initial Catalog=PBKS_WSTOCK;User Id=sa;Password=" & gPassword
        cn.Open strConnection
    End If
    
    Set oTF = New z_TextFile
    strLogPath = strSharedServerFolder & "\FETCHLOG" & Format(Date, "yyyymmdd") & ".txt"
    oTF.WriteToLogandsave strConnection, strLogPath
    Set FTP1 = New FTPClass
    Set oFSO = New FileSystemObject
    
    Set frmMain = New frmMain
    frmMain.Show
'    frmMain.Refresh
'    frmMain.DoWork
'     Unload frmMain
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Mainc.Main"
    HandleErrorQuiet True
End Sub

Public Function DownloadFolder() As String
    DownloadFolder = strSharedServerFolder & "\" & gDownloadFolder
End Function
Public Function Backupfolder() As String
    Backupfolder = strSharedServerFolder & "\" & gBackupFolder
End Function



Public Function Dialup() As Boolean
    Dialup = (gDialup = "TRUE")
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
    ErrorIn "PapyConn.NameOfPC"
End Property


