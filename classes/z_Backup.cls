VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim cmd As ADODB.Command
Dim Param As ADODB.Parameter
Dim ZIP
Dim strBUFolder As String
Dim fs As New FileSystemObject
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer
Event DiskBackupDone()
Event LocalBackupDone()

Private Sub ZipFileInPlace(pUnzippedFileName As String, ZippedFilename As String)
Dim bOK As Boolean
Dim strPos As String

    Set ZIP = CreateObject("FathZIP.FathZIPCtrl.1")
    ZIP.ProcessSubfolders = False
    ZIP.BasePath = ""
    ZIP.PreservePaths = False
    If fs.FileExists(ZippedFilename) Then
        fs.DeleteFile ZippedFilename, True
    End If
    
    ZIP.CreateZip ZippedFilename, ""
    If fs.FileExists(pUnzippedFileName) Then
        ZIP.AddFile pUnzippedFileName, ""
        bOK = (ZIP.LastError = 0)
    End If
    ZIP.Close
    If fs.FileExists(pUnzippedFileName) Then
        fs.DeleteFile pUnzippedFileName, True
    End If

    Set ZIP = Nothing

End Sub

'Public Function BackupToBriefcase(pFolder As String) As Boolean
'    On Error GoTo errHandler
'Dim strPos As String
'Dim strBackupFilename As String
'Dim OpenResult As Integer
'Dim oDMO As New z_SQLDMO
'Dim fs As New FileSystemObject
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandTimeout = 0
'    Set cmd.ActiveConnection = oPC.COShort
'    oDMO.ShrinkDatabase
'    If fs.FileExists(pFolder) Then fs.DeleteFile pFolder, True
'
'    strBackupFilename = fs.GetParentFolderName(pFolder) & "\" & fs.GetBaseName(pFolder) & ".BAK"
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "BACKUP DATABASE PBKS to disk = '" & strBackupFilename & "' WITH INIT, NAME = 'Full Backup of PBKSDATA'"
'    cmd.execute
'
'    If UCase(Right(pFolder, 3)) = "ZIP" Then
'        ZipFileInPlace strBackupFilename, fs.GetParentFolderName(pFolder) & "\" & fs.GetBaseName(pFolder) & ".ZIP"
'    End If
'    Set cmd = Nothing
'    BackupToBriefcase = True
'
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_PBKSBackup.BackupToBriefcase(pFolder)", pFolder
'End Function


Public Function BackupToLocal() As Boolean
Dim strPos As String
Dim OpenResult As Integer
    
    On Error GoTo ErrHandler
    
    If fs.FolderExists(oPC.BackupFolder) Then
        strBUFolder = oPC.BackupFolder
    Else
        strBUFolder = oPC.SharedFolderRoot & "\"
    End If
strPos = "6"

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = oPC.COShort
'    oDatabase.Shrink 10, SQLDMOShrink_Default
strPos = "7"
    If fs.FileExists(strBUFolder & "PBKS.BAK") Then fs.DeleteFile strBUFolder & "PBKS.BAK", True
    If fs.FileExists(strBUFolder & "PBKSMASTER.BAK") Then fs.DeleteFile strBUFolder & "PBKSMASTER.BAK", True

    cmd.CommandType = adCmdText
    cmd.CommandText = "BACKUP DATABASE PBKSC to disk = '" & strBUFolder & "PBKSC.BAK' WITH INIT, NAME = 'Full Backup of PBKSCDATA'"
    cmd.execute
    cmd.CommandText = "BACKUP DATABASE MASTER to disk = '" & strBUFolder & "PBKSMASTER.BAK' WITH INIT, NAME = 'Full Backup of PBKSMASTER'"
    cmd.execute

    Set cmd = Nothing
    BackupToLocal = True
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.BackupToLocal", , , , "position,BackupFolder", Array(strPos, strBUFolder)
End Function

Public Function ZIPBackupToLocal() As Boolean
    On Error GoTo ErrHandler
Dim bOK As Boolean
Dim strPos As String

strPos = "pos 1"
    Set ZIP = CreateObject("FathZIP.FathZIPCtrl.1")
    ZIP.ProcessSubfolders = False
    ZIP.BasePath = ""
    ZIP.PreservePaths = False
strPos = "pos 2"
    If fs.FileExists(oPC.BackupFolder & "PBKS.ZIP") Then
        fs.DeleteFile oPC.BackupFolder & "PBKS.ZIP", True
    End If
strPos = "pos 3"
    ZIP.CreateZip oPC.BackupFolder & "PBKS.ZIP", ""
    ZIP.AddFile oPC.BackupFolder & "PBKS.BAK", ""
strPos = "pos 4"
    bOK = (ZIP.LastError = 0)
    
    If bOK Then
        strPos = "pos 5"
        ZIP.AddFile oPC.BackupFolder & "PBKSMASTER.BAK", ""
        bOK = IIf(ZIP.LastError = 0, bOK, False)
    End If
strPos = "pos 6"
    ZIP.Close
strPos = "pos 7"
    Set ZIP = Nothing
    ZIPBackupToLocal = bOK
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.ZIPBackupToLocal:BU folder = " & oPC.BackupFolder, , , , "strPos", , Array(strPos)
End Function
Public Function ZIPBackupToNonLocal() As Boolean
    On Error GoTo ErrHandler
Dim bOK As Boolean
Dim strPos As String

    strPos = "CreateObject FathZip"
    Set ZIP = CreateObject("FathZIP.FathZIPCtrl.1")
    ZIP.ProcessSubfolders = False
    ZIP.BasePath = ""
    ZIP.PreservePaths = False
    strPos = "fs.FileExists " & oPC.RemovableBackupShareName & "*.*"
    If fs.FolderExists(oPC.RemovableBackupShareName) Then
        fs.DeleteFile oPC.RemovableBackupShareName & "*.*", True
    End If
    strPos = "zip.CreateZip " & oPC.RemovableBackupShareName & "PBKS.ZIP"
    ZIP.CreateZip oPC.RemovableBackupShareName & "PBKS.ZIP", ""
    If fs.FileExists(strBUFolder & "PBKSC.BAK") Then
        strPos = "zip.AddFile " & strBUFolder & "PBKSC.BAK"
        ZIP.AddFile strBUFolder & "PBKSC.BAK", ""
        bOK = (ZIP.LastError = 0)
    End If
    ZIP.Close

    Set ZIP = Nothing
    ZIPBackupToNonLocal = bOK
    RaiseEvent DiskBackupDone
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.ZIPBackupToNonLocal", , , , "StrPos = ", Array(strPos)
End Function
Public Function CopyBackupToNonLocal() As Boolean
    On Error GoTo ErrHandler
Dim bOK As Boolean
Dim strPos As String

    If fs.FolderExists(oPC.RemovableBackupShareName) Then
        fs.DeleteFile oPC.RemovableBackupShareName & "*.*", True
    End If
    If fs.FileExists(strBUFolder & "PBKS.BAK") Then
        fs.MoveFile strBUFolder & "PBKS.BAK", oPC.RemovableBackupShareName
    End If
    CopyBackupToNonLocal = bOK
    
    RaiseEvent DiskBackupDone
    Exit Function
ErrHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.CopyBackupToNonLocal", , , , "StrPos = ", Array(strPos)
End Function


