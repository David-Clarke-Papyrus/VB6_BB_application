VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_MDBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public DBConn As ADODB.Connection

Dim sDBPath As String
Dim strServerRootPath As String
Dim strSharedFolderRoot As String
Dim strBackupFolder As String
Dim strDDB As String
Dim bUsesWORD As Boolean
Dim strServer As String

Const DB_NAME = "PapyPos.mdb"

Public Function dbConnect() As Integer
    On Error GoTo errHandler

Dim szConnect As String
    dbConnect = 0
    SetServername
    If DBConn Is Nothing Then
        szConnect = "Provider=Microsoft.jet.OLEDB.4.0;DATA SOURCE=" & oPC.SharedFolderRoot & "\DATA\PapyPOS.MDB" _
                            & ";JET OLEDB:SYSTEM DATABASE=" & oPC.SharedFolderRoot & "\DATA\POS.MDW;USER ID=POS;Password=car;"
        Set DBConn = New ADODB.Connection

        DBConn.Open szConnect
    End If
    Set oPC = Me
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.dbConnect"
End Function
Public Function dbCloseConnect() As Integer
    On Error GoTo errHandler
    If Not DBConn Is Nothing Then
        If DBConn.State = adStateOpen Then DBConn.Close
        Set DBConn = Nothing
    End If
EXIT_Handler:
    Exit Function
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.dbCloseConnect"
End Function

Private Sub LoadDBPath()
    On Error GoTo errHandler
Dim fs As New Scripting.FileSystemObject

TryAgain:
    If sDBPath = "" Then
        'USER CANCELLED
        If MsgBox("Application won't open without valid Database Path!" & vbLf & _
                  "Try again?", vbYesNo + vbExclamation, "Select Database") = vbNo Then
            Exit Sub
        Else
            GoTo TryAgain
        End If
    
    End If
    sDBPath = IIf(Right(sDBPath, 1) = "\", sDBPath, sDBPath & "\")
    If Not fs.FileExists(sDBPath & DB_NAME) Then
        If MsgBox("The Database Path provided is not valid!" & vbLf & _
                  "The Application won't open without valid Database Path!" & vbLf & _
                  "Try again?", vbYesNo + vbExclamation, "Select Database") = vbNo Then
            sDBPath = ""
            Exit Sub
        Else
            GoTo TryAgain
        End If
    End If
    SaveSetting App.Title, "Settings", "DBPath", sDBPath
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.LoadDBPath"
End Sub
Friend Property Get DBPath() As String
    On Error GoTo errHandler
    DBPath = sDBPath
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.DBPath"
End Property
Friend Property Get DBName() As String
    On Error GoTo errHandler
    DBName = DB_NAME
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.DBName"
End Property

Private Sub SetServername()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim oTF As New z_TextFile
Dim strTag As String
Dim strTmp As String
Dim strValue As String

    strServerRootPath = fs.GetParentFolderName(App.Path)
'    MsgBox "ParentFolderName = " & strServerRootPath
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
    
    
    
'ERRH:
'    Exit Sub
'    Resume
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.SetServername"
End Sub
Public Property Get SharedFolderRoot() As String
    On Error GoTo errHandler
    SharedFolderRoot = strSharedFolderRoot
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.SharedFolderRoot"
End Property

Private Sub Class_Initialize()
    On Error GoTo errHandler
    SetServername
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.Class_Initialize"
End Sub
Public Property Get servername() As String
    On Error GoTo errHandler
    servername = strServer
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_MDBConnection.servername"
End Property

