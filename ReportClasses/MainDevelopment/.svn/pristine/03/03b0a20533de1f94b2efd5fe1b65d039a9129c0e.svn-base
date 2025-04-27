VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_PBKSBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cmd As ADODB.Command
Dim Param As ADODB.Parameter
Dim zip
Dim strBUFolder As String
Dim fs As New FileSystemObject
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer
Event DiskBackupDone()
Dim oFSO As New FileSystemObject

Public Function BackupToLocal() As Boolean
Dim strPos As String
    
    On Error GoTo errHandler
    
    If fs.FolderExists(oPC.BackupFolder) Then
        strBUFolder = oPC.BackupFolder
    Else
        strBUFolder = oPC.SharedFolderRoot
    End If
strPos = "6"

    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = cnPapy
strPos = "7"
    If fs.FileExists(strBUFolder & "PBKS.BAK") Then fs.DeleteFile strBUFolder & "PBKS.BAK", True
    If fs.FileExists(strBUFolder & "PBKSMASTER.BAK") Then fs.DeleteFile strBUFolder & "PBKSMASTER.BAK", True

    cmd.CommandType = adCmdText
    cmd.CommandText = "BACKUP DATABASE PBKS to disk = '" & strBUFolder & "PBKS.BAK' WITH INIT, NAME = 'Full Backup of PBKSDATA'"
    cmd.Execute
    cmd.CommandText = "BACKUP DATABASE MASTER to disk = '" & strBUFolder & "PBKSMASTER.BAK' WITH INIT, NAME = 'Full Backup of PBKSMASTER'"
    cmd.Execute

    Set cmd = Nothing
    BackupToLocal = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.BackupToLocal", , , , "position,BackupFolder", Array(strPos, strBUFolder)
End Function

Public Function ZIPBackupToLocal() As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean

    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    If fs.FileExists(oPC.BackupFolder & "PBKS.ZIP") Then
        fs.DeleteFile oPC.BackupFolder & "PBKS.ZIP", True
    End If
    zip.CreateZip oPC.BackupFolder & "PBKS.ZIP", ""

    zip.AddFile oPC.BackupFolder & "PBKS.BAK", ""
    bOK = (zip.LastError = 0)
    If bOK Then
        zip.AddFile oPC.BackupFolder & "PBKSMASTER.BAK", ""
        bOK = IIf(zip.LastError = 0, bOK, False)
    End If

    zip.Close

    Set zip = Nothing
    ZIPBackupToLocal = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.ZIPBackupToLocal"
End Function
Public Function ZIPBackupToNonLocal() As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim strPos As String

    strPos = "CreateObject FathZip"
    Set zip = CreateObject("FathZIP.FathZIPCtrl.1")
    strPos = "fs.FileExists " & oPC.RemovableBackupShareName & "*.*"
    If fs.FileExists(oPC.RemovableBackupShareName & "*.*") Then
        strPos = "fs.DeleteFile " & oPC.RemovableBackupShareName & "*.*"
        fs.DeleteFile oPC.RemovableBackupShareName & "*.*", True
    End If
    strPos = "zip.CreateZip " & oPC.RemovableBackupShareName & "PBKS.ZIP"
    If oFSO.FolderExists(oPC.RemovableBackupShareName) Then
        zip.CreateZip oPC.RemovableBackupShareName & "PBKS.ZIP", ""
        If fs.FileExists(strBUFolder & "PBKS.BAK") Then
            strPos = "zip.AddFile " & strBUFolder & "PBKS.BAK"
            zip.AddFile strBUFolder & "PBKS.BAK", ""
            bOK = (zip.LastError = 0)
        End If
        zip.Close
    Else
        ZIPBackupToNonLocal = False
        Exit Function
    End If
    Set zip = Nothing
    ZIPBackupToNonLocal = bOK
    RaiseEvent DiskBackupDone
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.ZIPBackupToNonLocal", , , , strPos, Array(strPos)
End Function
