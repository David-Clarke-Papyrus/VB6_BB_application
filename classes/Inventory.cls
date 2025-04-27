VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Inventory"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strPCName As String
Dim strLocalRootFolder As String
Dim strSharedFolderRoot As String
Dim strCENTRALSERVERMACHINE As String
Dim strPassword As String
Dim strServerName As String
Dim strMainConnectionString As String
Dim zBF As New zc_BF
Dim oBookfindmanager As z_BookfindManager
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
Public Function DoBookfind2() As String
    DoBookfind2 = "In Process"
End Function

Public Function DoBookfind() As String
    strPCName = Trim(Me.NameOfPC)

    If IsNetConnectionAlive Then
        strLocalRootFolder = "\\" & strPCName & "\PBKS_S"
        strCENTRALSERVERMACHINE = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "NETWORK", "CENTRALSERVERMACHINE", "")
        strSharedFolderRoot = "\\" & strCENTRALSERVERMACHINE & "\PBKS_S"
    Else
        strLocalRootFolder = "C:\PBKS"
        strSharedFolderRoot = strLocalRootFolder
    End If

    strPassword = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "NETWORK", "PASSWORD", "")
    strServerName = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "NETWORK", "MAINSQLSERVER", "")
    strBookfindRoot = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "BOOKDATA", "BOOKFINDROOT", "")
    strBookfindFacet = GetIniKeyValue(strLocalRootFolder & "\CENTRAL.INI", "BOOKDATA", "BOOKFINDFACET", "WEBK")
    strMainConnectionString = "Provider=SQLOLEDB;Persist Security Info=False;User ID=sa;Data Source=" & strServerName & ";Initial Catalog=PBKSC;User Id=sa;Password=" & strPassword & ";Connect Timeout=120"
    bBOOKFINDISBN13ENABLED = GetIniKeyValue(strSharedFolderRoot & "\CENTRAL.INI", "BOOKDATA", "BOOKFINDISBN13ENABLED", "FALSE") = "TRUE"
    Set cnPapyShort = New ADODB.connection
    If cnPapyShort.errors.Count > 0 Or cnPapyShort.State = adStateClosed Then
        On Error Resume Next
        cnPapyShort.Close
        cnPapyShort.Open strMainConnectionString
        cnPapyShort.CommandTimeout = 360
    End If

    Set oBookfindmanager = New z_BookfindManager
    oBookfindmanager.PrepareBookfind
    zBF.UpdateFromBookfind strBookfindFacet
    cnPapyShort.Close
    Set cnPapyShort = Nothing
End Function



Public Property Get NameOfPC() As String
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
End Property

