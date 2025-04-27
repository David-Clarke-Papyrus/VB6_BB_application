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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cmd As ADODB.Command
Dim Param As ADODB.Parameter
Dim Zip
Dim strBUFolder As String
Dim fs As New FileSystemObject
'Dim oDatabase As SQLDMO.Database2
'Dim oSQLServer As SQLDMO.SQLServer
Event DiskBackupDone()
Event LocalBackupDone()

Private Sub ZipFileInPlace(pUnzippedFileName As String, ZippedFilename As String)
Dim bOK As Boolean
Dim strPos As String

    Set Zip = CreateObject("FathZIP.FathZIPCtrl.1")
    Zip.ProcessSubfolders = False
    Zip.BasePath = ""
    Zip.PreservePaths = False
    If fs.FileExists(ZippedFilename) Then
        fs.DeleteFile ZippedFilename, True
    End If
    
    Zip.CreateZip ZippedFilename, ""
    If fs.FileExists(pUnzippedFileName) Then
        Zip.AddFile pUnzippedFileName, ""
        bOK = (Zip.LastError = 0)
    End If
    Zip.Close
    If fs.FileExists(pUnzippedFileName) Then
        fs.DeleteFile pUnzippedFileName, True
    End If

    Set Zip = Nothing

End Sub

Public Function BackupToBriefcase(pFolder As String, Optional RestoreLocallyAsTest As Boolean = False, Optional BackupCurrentlyConnectedDB As Boolean = False) As Boolean
    On Error GoTo errHandler
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
'    cmd.CommandText = "BACKUP DATABASE " & IIf(BackupCurrentlyConnectedDB, oPC.DatabaseName, "PBKS") & " to disk = '" & strBackupFilename & "' WITH INIT, NAME = 'Full Backup of PBKSDATA'"
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
'    If RestoreLocallyAsTest Then
'        oDMO.RestoreDatabase pFolder
'    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.BackupToBriefcase(pFolder)", pFolder
End Function


Public Function BackupToLocal() As Boolean
Dim strPos As String
Dim OpenResult As Integer
    
    On Error GoTo errHandler
    
'    If fs.FolderExists(oPC.BackupFolder) Then
'        strBUFolder = oPC.BackupFolder
'    Else
'        strBUFolder = oPC.SharedFolderRoot & "\"
'    End If
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandTimeout = 0
'    Set cmd.ActiveConnection = oPC.COShort
'    If fs.FileExists(strBUFolder & "PBKS.BAK") Then fs.DeleteFile strBUFolder & "PBKS.BAK", True
'    If fs.FileExists(strBUFolder & "PBKSMASTER.BAK") Then fs.DeleteFile strBUFolder & "PBKSMASTER.BAK", True
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "BACKUP DATABASE " & oPC.DatabaseName & " to disk = '" & strBUFolder & "PBKS.BAK' WITH INIT, NAME = 'Full Backup of PBKSDATA'"
'    cmd.execute
'    cmd.CommandText = "BACKUP DATABASE MASTER to disk = '" & strBUFolder & "PBKSMASTER.BAK' WITH INIT, NAME = 'Full Backup of PBKSMASTER'"
'    cmd.execute
'
'    Set cmd = Nothing
'    BackupToLocal = True
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.BackupToLocal", , , , "position,BackupFolder", Array(strPos, strBUFolder)
End Function

Public Function ZIPBackupToLocal() As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim strPos As String

strPos = "pos 1"
    Set Zip = CreateObject("FathZIP.FathZIPCtrl.1")
    Zip.ProcessSubfolders = False
    Zip.BasePath = ""
    Zip.PreservePaths = False
strPos = "pos 2"
    If fs.FileExists(oPC.BackupFolder & "PBKS.ZIP") Then
        fs.DeleteFile oPC.BackupFolder & "PBKS.ZIP", True
    End If
strPos = "pos 3"
    Zip.CreateZip oPC.BackupFolder & "PBKS.ZIP", ""
    Zip.AddFile oPC.BackupFolder & "PBKS.BAK", ""
strPos = "pos 4"
    bOK = (Zip.LastError = 0)
    
    If bOK Then
        strPos = "pos 5"
        Zip.AddFile oPC.BackupFolder & "PBKSMASTER.BAK", ""
        bOK = IIf(Zip.LastError = 0, bOK, False)
    End If
strPos = "pos 6"
    Zip.Close
strPos = "pos 7"
    Set Zip = Nothing
    ZIPBackupToLocal = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.ZIPBackupToLocal:BU folder = " & oPC.BackupFolder, , , , "strPos", , Array(strPos)
End Function
Public Function ZIPBackupToNonLocal() As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim strPos As String
Dim f As Folder
Dim fil As File
Dim fc

    strPos = "CreateObject FathZip"
    Set Zip = CreateObject("FathZIP.FathZIPCtrl.1")
    Zip.ProcessSubfolders = False
    Zip.BasePath = ""
    Zip.PreservePaths = False
    strPos = "fs.FileExists " & oPC.RemovableBackupShareName & "*.*"
    Set f = fs.GetFolder(oPC.RemovableBackupShareName)
    Set fc = f.Files
    For Each fil In fc
        If fil.Name = "PBKS.ZIP" Then
            fil.Delete True
        End If
    Next
    
    
    strPos = "zip.CreateZip " & oPC.RemovableBackupShareName & "PBKS.ZIP"
    Zip.CreateZip oPC.RemovableBackupShareName & "PBKS.ZIP", ""
    If fs.FileExists(strBUFolder & "PBKS.BAK") Then
        strPos = "zip.AddFile " & strBUFolder & "PBKS.BAK"
        Zip.AddFile strBUFolder & "PBKS.BAK", ""
        bOK = (Zip.LastError = 0)
    End If
    Zip.Close

    Set Zip = Nothing
    ZIPBackupToNonLocal = bOK
    RaiseEvent DiskBackupDone
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.ZIPBackupToNonLocal", , , , "StrPos = ", Array(strPos)
End Function
Public Function CopyBackupToNonLocal() As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim strPos As String

                    strPos = "fs.FileExists " & oPC.RemovableBackupShareName & "*.*"
    If fs.FileExists(oPC.RemovableBackupShareName & "*.*") Then
                    strPos = "fs.DeleteFile " & oPC.RemovableBackupShareName & "*.*"
        fs.DeleteFile oPC.RemovableBackupShareName & "*.*", True
    End If
                    strPos = "zip.CreateZip " & oPC.RemovableBackupShareName & "PBKS.ZIP"
    If fs.FileExists(strBUFolder & "PBKS.BAK") Then
                    strPos = "zip.AddFile " & strBUFolder & "PBKS.BAK"
        fs.MoveFile strBUFolder & "PBKS.BAK", oPC.RemovableBackupShareName
    End If
    CopyBackupToNonLocal = bOK
    
    RaiseEvent DiskBackupDone
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.CopyBackupToNonLocal", , , , "StrPos = ", Array(strPos)
End Function

Public Function BackupFDToLocal() As Boolean
Dim strPos As String
Dim OpenResult As Integer
    
    On Error GoTo errHandler
    
'    If fs.FolderExists(oPC.BackupFolder) Then
'        strBUFolder = oPC.BackupFolder
'    Else
'        strBUFolder = oPC.SharedFolderRoot & "\"
'    End If
'strPos = "6"
'
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = New ADODB.Command
'    cmd.CommandTimeout = 0
'    Set cmd.ActiveConnection = oPC.COShort
'    If fs.FileExists(strBUFolder & "PBKSFD.BAK") Then fs.DeleteFile strBUFolder & "PBKSFD.BAK", True
'
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "BACKUP DATABASE PBKSFD to disk = '" & strBUFolder & "PBKSFD.BAK' WITH INIT, NAME = 'Full Backup of PBKSFDDATA'"
'    cmd.execute
'
'    Set cmd = Nothing
'    BackupFDToLocal = True
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_PBKSBackup.BackupFDToLocal", , , , "position,BackupFolder", Array(strPos, strBUFolder)
End Function

