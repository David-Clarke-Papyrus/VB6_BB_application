Attribute VB_Name = "Default"
Option Explicit

Global cn As ADODB.Connection
Global gLocalRootFolder As String
Global oTF As New z_TextFile
Global oFSO As FileSystemObject
Global strServerName As String
Global gRes

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
Dim strPCName As String
Global fINET As wininet
Global gConnectionName As String
Global gPassword As String
Global gPCNAME As String

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Private Sub Main()
    On Error GoTo errHandler
Dim frmMain As frmMain
    
    If App.PrevInstance Then
       ActivatePrevInstance
       Exit Sub
    End If
    gLocalRootFolder = "C:\PBKS"
   
    strServerName = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "NETWORK", "SQLSERVERNAME", "")
    gFTPSourceAddress = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "FTPADDRESS", "")
    gFTPSourceUsername = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "FTPUSERNAME", "")
    gFTPSourcePassword = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "FTPPASSWORD", "")
    gFTPSourceFolder = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "FTPFOLDER", "")
    gFTPTargetAddress = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "CENTRAL", "FTPADDRESS", "")
    gFTPTargetUsername = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "CENTRAL", "FTPUSERNAME", "")
    gFTPTargetPassword = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "CENTRAL", "FTPPASSWORD", "")
    gFTPTargetFolder = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "CENTRAL", "FTPFOLDER", "")
    gDownloadFolder = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "DOWNLOADFOLDER", "")
    gBackupFolder = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "BACKUPFOLDER", "BU")
    gStoreCode = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "NETWORK", "STORECODE", "")
    gDialup = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "INTERNETDIALUP", "TRUE")
    gConnectionName = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "SUPPORT", "INTERNETDIALUP", "TRUE")
    gPassword = GetIniKeyValue(gLocalRootFolder & "\PBKS_WSTOCK.INI", "NETWORK", "PASSWORD", "")
    If cn Is Nothing Then
        Set cn = New ADODB.Connection
        cn.Provider = "sqloledb"
        cn.ConnectionTimeout = 60
        cn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;Data Source=" & strServerName & ";Initial Catalog=PBKS_WSTOCK;User Id=sa;Password=" & gPassword
    End If
    gPCNAME = Trim(NameOfPC)
    Set oTF = New z_TextFile
    oTF.OpenTextFile gLocalRootFolder & "\FETCHLOG" & Format(Date, "yyyymmdd") & ".txt"
    Set FTP1 = New FTPClass
    Set oFSO = New FileSystemObject
    
    Set frmMain = New frmMain
    frmMain.Show vbModal
   ' frmMain.Refresh
   ' frmMain.DoWork
    Unload frmMain
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Mainc.Main"
    HandleErrorQuiet True
End Sub
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

Public Function DownloadFolder() As String
    DownloadFolder = gLocalRootFolder & "\" & gDownloadFolder
End Function
Public Function Backupfolder() As String
    Backupfolder = gLocalRootFolder & "\" & gBackupFolder
End Function



Public Function Dialup() As Boolean
    Dialup = (gDialup = "TRUE")
End Function


