VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_POSCLIConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mCurrentSecurityCode As String
Public DBConn As ADODB.Connection
Dim oGD As z_GetData
Dim sDBPath As String
Dim strServerRootPath As String
Dim strSharedFolderRoot As String
Dim strBackupFolder As String
Dim strDDB As String
Dim bUsesWORD As Boolean
Dim strServer As String

Const DB_NAME = "PapyPos.mdb"
Public Function dbConnect()
    On Error GoTo errHandler

    If DBConn Is Nothing Then
        Set DBConn = New ADODB.Connection
        DBConn.Provider = "sqloledb"
        DBConn.Open "Data Source=" & Me.servername & ";Initial Catalog=PBKSFD;User Id=sa" & ";Password=" & "" & "; Network Library=dbmssocn;"
        DBConn.CommandTimeout = 0
    Else
    End If
    Set oGD = New z_GetData
    oGD.LoadSPers
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.dbConnect"
End Function

Public Function dbCloseConnect() As Integer
    On Error GoTo errHandler
    If Not DBConn Is Nothing Then
        If DBConn.State = adStateOpen Then DBConn.Close
        Set DBConn = Nothing
    End If
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.dbCloseConnect"
End Function

Friend Property Get DBPath() As String
    On Error GoTo errHandler
    DBPath = sDBPath
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.DBPath"
End Property
Friend Property Get DBName() As String
    On Error GoTo errHandler
    DBName = DB_NAME
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.DBName"
End Property

Private Sub SetServername()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim oTF As New z_TextFile
Dim strTag As String
Dim strTmp As String
Dim strValue As String

    strServerRootPath = fs.GetParentFolderName(App.Path)
    oTF.OpenTextFileToRead strServerRootPath & "\PBKS.INI"                        'ServerRootPath & "ServerInstance.txt"
    strServer = oTF.ReadLinefromTextFile

    Do While Not oTF.IsEOF
        strTmp = oTF.ReadLinefromTextFile
        strTag = Replace(Left(strTmp, InStr(strTmp, "}") - 1), "{", "")
        strValue = Right(strTmp, Len(strTmp) - InStr(1, strTmp, "="))
        Select Case UCase(strTag)
        Case "USESWORD"
            bUsesWORD = IIf(strValue = "TRUE", True, False)
        Case "SHAREDFOLDER"
            strSharedFolderRoot = strValue
        Case "BACKUPFOLDER"
            strBackupFolder = strValue
        Case "BFDISTRIBUTORDB"
            strDDB = strValue
        End Select
    Loop
    oTF.CloseTextFile
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.SetServername"
End Sub
Public Property Get SharedFolderRoot() As String
    On Error GoTo errHandler
    SharedFolderRoot = strSharedFolderRoot
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.SharedFolderRoot"
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
    SetServername
    Set oPC = Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.Class_Initialize"
End Sub
Public Property Get servername() As String
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim oTF As New z_TextFile
Dim strTag As String
Dim strTmp As String
Dim strValue As String

    strServerRootPath = fs.GetParentFolderName(App.Path)
    oTF.OpenTextFileToRead strServerRootPath & "\PBKS.INI"                        'ServerRootPath & "ServerInstance.txt"
    strServer = oTF.ReadLinefromTextFile
    servername = strServer
    Do While Not oTF.IsEOF
        strTmp = oTF.ReadLinefromTextFile
        strTag = Replace(Left(strTmp, InStr(strTmp, "}") - 1), "{", "")
        strValue = Right(strTmp, Len(strTmp) - InStr(1, strTmp, "="))
        Select Case UCase(strTag)
        Case "USESWORD"
            bUsesWORD = IIf(strValue = "TRUE", True, False)
        Case "SHAREDFOLDER"
            strSharedFolderRoot = strValue
        Case "BACKUPFOLDER"
            strBackupFolder = strValue
        End Select
    Loop
    oTF.CloseTextFile
'ERRH:
'    Exit Property
'    Resume
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.Servername"
End Property

Public Property Let CurrentSecurityCode(val As String)
    mCurrentSecurityCode = val
End Property
Public Property Get CurrentSecurityCode() As String
    CurrentSecurityCode = mCurrentSecurityCode
End Property

Public Property Get GD() As z_GetData
    Set GD = oGD
End Property
