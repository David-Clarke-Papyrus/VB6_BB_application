VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Dayend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Dim WithEvents oCD As Z_CDWrite
Event STATUS(pMsg As String, pErr As Boolean)
Event Complete()
Event BackupComplete()
Dim bDoDayend  As Boolean
Dim dteDate As Date
Dim WithEvents oBatch As z_Batch
Attribute oBatch.VB_VarHelpID = -1
Dim WithEvents oBU As z_PBKSBackup
Attribute oBU.VB_VarHelpID = -1
Dim gSTAFFID As Long

Private Sub oBatch_DEResult(pMsg As String, pStatus As Long)
    On Error GoTo errHandler
    RaiseEvent STATUS(pMsg, (pStatus = 99))
    
    Exit Sub
errHandler:
    ErrorIn "z_Dayend.oBatch_DEResult(pMsg,pStatus)", Array(pMsg, pStatus)
End Sub

Private Sub oBU_DiskBackupDone()
    On Error GoTo errHandler
        
        RaiseEvent BackupComplete
        
        If bDoDayend Then DayendUpdate gSTAFFID
        
    Exit Sub
errHandler:
    ErrorIn "z_Dayend.oBU_DiskBackupDone"
End Sub

'Private Sub oCD_Complete()
'    On Error GoTo errHandler
'
'        RaiseEvent BackupComplete
'
'        If bDoDayend Then DayendUpdate gSTAFFID
'
'    Exit Sub
'errHandler:
'    ErrorIn "z_Dayend.oCD_Complete"
'End Sub
'
'
'Private Sub oCD_Status(pMsg As String, bErr As Boolean)
'    On Error GoTo errHandler
'
'        If bErr Then
'            If bDoDayend Then
'                oPC.CO.Execute "INSERT INTO tOPERATION (OP_STARTEDAT,OP_TYPE,OP_RESULT) VALUES(GETDATE(),1,2)"
'            End If
'        End If
'        RaiseEvent Status(pMsg, bErr)
'
'    Exit Sub
'errHandler:
'    ErrorIn "z_Dayend.oCD_Status(pMsg,bErr)", Array(pMsg, bErr)
'End Sub
'
'Private Sub oCD_EraseDiscComplete()
'    On Error GoTo errHandler
'    WriteBackup
'    Exit Sub
'errHandler:
'    ErrorIn "z_Dayend.oCD_EraseDiscComplete"
'End Sub
'
'Private Sub oCD_EraseDiscStart()
'    On Error GoTo errHandler
'    RaiseEvent Status("Erasing CD prior to writing . . .", False)
'    Exit Sub
'errHandler:
'    ErrorIn "z_Dayend.oCD_EraseDiscStart"
'End Sub
'

'Private Sub WriteBackup()
'    On Error GoTo errHandler
'Dim strErrmsg As String
'    If oPC.CompressBackup = True Then
'        oCD.AddFile oPC.BackupFolder & "PBKS.ZIP"
'    Else
'        oCD.AddFile oPC.BackupFolder & "PBKS.BAK"
'        oCD.AddFile oPC.BackupFolder & "PBKSMASTER.BAK"
'    End If
'    Exit Sub
'errHandler:
'    ErrorIn "z_Dayend.WriteBackup"
'End Sub

Public Sub RunDayend(pSTAFFID As Long)
    bDoDayend = True
    gSTAFFID = pSTAFFID
    Backup
End Sub
Public Sub RunBackupOnly()
    bDoDayend = False
    Backup
End Sub

Public Sub Backup()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
                    p 1
    Set oBU = New z_PBKSBackup
    RaiseEvent STATUS("Backing up to local disk", False)
    DoEvents
    Check oBU.BackupToLocal, EXC_GENERAL, "Backup was not successful.Contact support person"
    Check (fs.FileExists(oPC.BackupFolder & "PBKS.BAK") And fs.FileExists(oPC.BackupFolder & "PBKSMASTER.BAK")), EXC_GENERAL, "Backup was not successful.Contact support person"
                    p 2
    If oPC.RemoveableBackupMedium = "CD" Then
                    p 3
'        If oPC.CompressBackup = True Then
'            Check oBU.ZIPBackupToLocal, EXC_GENERAL, "Backup zipping operation to removeable storage was not successful. Contact support person"
'        End If
'        Set oCD = New Z_CDWrite
'        Check oCD.SetDefaultDrive(""), EXC_GENERAL, "No CD drive to write to."
'        If oCD.ISRewritable Then
'                    p 4
'            oCD.EraseCD
'        Else
'                    p 5
'            WriteBackup
'        End If
    Else
                    p 6
        If oPC.CompressBackup = True Then
                    p 7
            RaiseEvent STATUS("Compressing file to removable disk", False)
            DoEvents
            Check oBU.ZIPBackupToNonLocal, EXC_GENERAL, "Fail backing up to Non-Local"
        Else
            oBU.CopyBackupToNonLocal
        End If
    End If
    Set oBU = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.Backup", , , , "Position", Array(strErrPos)
End Sub


Public Sub DayendUpdate(gSTAFFID As Long)
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim oDMO As New z_SQLDMO
        
    Set oBatch = New z_Batch
    RaiseEvent STATUS("Running dayend procedure", False)
    lngResult = oBatch.DailyUpdate(Date, gSTAFFID)
    If lngResult = 0 Then
        RaiseEvent STATUS("Rebuilding indexes", False)
        oSQL.RunProc "RebuildAllIndexes", Array(), ""
        
        RaiseEvent STATUS("Removing old files in temporary folders", False)
        CleanupOldFilesInTempFolders
        
    'If POS is not active we must ensure that the usual POS triggers are disabled
        RaiseEvent STATUS("Checking triggers", False)
        oDMO.SwitchPOSTriggers (oPC.POSActive)
        
        Set oBatch = Nothing
        
        RaiseEvent STATUS("Dayend complete - CHECK BACKUP DEVICE CONTAINS BACKUP FILE", False)
    End If
    
    RaiseEvent Complete

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.DayendUpdate"
End Sub


Private Sub CleanupOldFilesInTempFolders()
Dim fs As New FileSystemObject
Dim fol As Object
Dim fils As Object
Dim f As File
    
    On Error Resume Next
    
    If fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        Set fils = fs.GetFolder(oPC.SharedFolderRoot & "\TEMP").Files
        For Each f In fils
            If DateDiff("d", f.DateCreated, Date) > oPC.KeepTemporaryfilesFor_n_Days Then
                f.Delete
            End If
        Next
    End If
    If fs.FolderExists(oPC.SharedFolderRoot & "\PDF") Then
        Set fils = fs.GetFolder(oPC.SharedFolderRoot & "\PDF").Files
        For Each f In fils
            If DateDiff("d", f.DateCreated, Date) > oPC.KeepTemporaryfilesFor_n_Days Then
                f.Delete
            End If
        Next
    End If

End Sub

