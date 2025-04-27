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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents oCD As Z_CDWrite
Attribute oCD.VB_VarHelpID = -1
Event Status(pMsg As String)
Event BackupComplete()
Dim bDoDayend  As Boolean
Dim dteDate As Date
Dim oB As z_Batch
Dim WithEvents oBU As z_PBKSBackup
Attribute oBU.VB_VarHelpID = -1
Dim lngResult As Long
Private Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long

Private Sub Class_Initialize()
    On Error GoTo errHandler
    strLocalRootFolder = "\\" & NameOfPC & "\PBKS_S"
    strSQLServerName = GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "NETWORK", "MAINSQLSERVER", "")
    strServerMachineName = GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "NETWORK", "PBKSSERVERMACHINE", "")
    strBackupMachineName = GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "NETWORK", "BACKUPMACHINE", "")
    strCOMPRESSBACKUP = GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "OPTIONS", "BACKUPCOMPRESSION ", "YES")
    strSharedServerFolder = "\\" & strServerMachineName & "\PBKS_S"
    strSourcePath = strSharedServerFolder & "\JOBS"
    strLocalBackupFolder = strSharedServerFolder & "\BU"
    strRemovableSharename = GetIniKeyValue(strLocalRootFolder & "\PBKS.INI", "FOLDERS", "BACKUP", "")
    strRemovable = strBackupMachineName & "\" & strRemovableSharename
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.Class_Initialize", , EA_NORERAISE
    HandleError
End Sub
Public Function NameOfPC() As String
    On Error GoTo errHandler
Dim NameSize As Long
Dim MachineName As String * 16
Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    NameOfPC = Left(MachineName, NameSize)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.NameOfPC"
End Function

Private Sub oBU_DiskBackupDone()
    On Error GoTo errHandler
'For backup to removable disk
Dim oDMO As New z_SQLDMO
    If bDoDayend Then
        Set oB = New z_Batch
        RaiseEvent Status("Running dayend procedure")
        lngResult = oB.DailyUpdate(dteDate)
        RaiseEvent Status("Rebuilding indexes")
        oDMO.RebuildIndexes
        RaiseEvent Status("Shrinking database")
        oDMO.ShrinkDatabase
        Set oB = Nothing
        RaiseEvent Status("Backup and dayend complete")
    Else
        RaiseEvent BackupComplete
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.oBU_DiskBackupDone", , EA_NORERAISE
    HandleError
End Sub
Private Sub oCD_Complete()
    On Error GoTo errHandler
'For backup to CD
Dim oDMO As New z_SQLDMO
    If bDoDayend Then
        Set oB = New z_Batch
        RaiseEvent Status("Running dayend procedure")
        lngResult = oB.DailyUpdate(dteDate)
        RaiseEvent Status("Rebuilding indexes")
        oDMO.RebuildIndexes
        RaiseEvent Status("Shrinking database")
        oDMO.ShrinkDatabase
        Set oB = Nothing
        RaiseEvent Status("Backup and dayend complete")
    Else
        RaiseEvent BackupComplete
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.oCD_Complete", , EA_NORERAISE
    HandleError
End Sub

Private Sub oB_DEResult(pMsg As String, pStatus As Long)
    On Error GoTo errHandler
    RaiseEvent Status(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.oB_DEResult(pMsg,pStatus)", Array(pMsg, pStatus), EA_NORERAISE
    HandleError
End Sub

Private Sub oCD_Status(pMsg As String, bErr As Boolean)
    On Error GoTo errHandler
    RaiseEvent Status(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.oCD_Status(pMsg,bErr)", Array(pMsg, bErr), EA_NORERAISE
    HandleError
End Sub

Private Sub oCD_EraseDiscComplete()
    On Error GoTo errHandler
    WriteBackup
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.oCD_EraseDiscComplete", , EA_NORERAISE
    HandleError
End Sub

Private Sub oCD_EraseDiscStart()
    On Error GoTo errHandler
    RaiseEvent Status("Erasing CD prior to writing . . .")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.oCD_EraseDiscStart", , EA_NORERAISE
    HandleError
End Sub


Private Sub WriteBackup()
    On Error GoTo errHandler
Dim strErrmsg As String
    If strCOMPRESSBACKUP = "YES" Then
        oCD.AddFile strLocalBackupFolder & "PBKS.ZIP"
        Check oCD.WriteDisc(strErrmsg), EXC_GENERAL, strErrmsg
    Else
        oCD.AddFile strLocalBackupFolder & "PBKS.BAK"
        oCD.AddFile strLocalBackupFolder & "PBKSMASTER.BAK"
        Check oCD.WriteDisc(strErrmsg), EXC_GENERAL, strErrmsg
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.WriteBackup"
End Sub

Public Sub Backup()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
    Set oBU = New z_PBKSBackup
    RaiseEvent Status("Backing up to local disk")
    DoEvents
   ' MsgBox "Code missing here"
    Check oBU.BackupToLocal, EXC_GENERAL, "Backup was not successful.Contact support person"
    Check (fs.FileExists(strLocalBackupFolder & "PBKS.BAK") And fs.FileExists(strLocalBackupFolder & "PBKSMASTER.BAK")), EXC_GENERAL, "Backup was not successful.Contact support person"
    
    If strREMOVABLEBACKUPTYPE = "CD" Then
        If strCOMPRESSBACKUP = "YES" Then
            Check oBU.ZIPBackupToLocal, EXC_GENERAL, "Backup zipping operation to removeable storage was not successful. Contact support person"
            Set oCD = New Z_CDWrite
            Check oCD.SetDefaultDrive(""), EXC_GENERAL, "No CD drive to write to."
            oCD.EraseCD
        Else
            Set oCD = New Z_CDWrite
            Check oCD.SetDefaultDrive(""), EXC_GENERAL, "No CD drive to write to"
            oCD.EraseCD
        End If
    Else
        If strCOMPRESSBACKUP = "YES" Then
            RaiseEvent Status("Compressing file to removable disk")
            DoEvents
            If Not oBU.ZIPBackupToNonLocal Then   ', EXC_GENERAL, "Fail backing up to Non-Local"
                RaiseEvent Status("Could not backup. Dayend not done.")
            End If
        End If
    End If
    Set oBU = Nothing
    
    Exit Sub
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.Backup"
End Sub

Public Sub DailyUpdateincBackup(pDate As Date)
    On Error GoTo errHandler
    bDoDayend = True
    dteDate = pDate
    Backup
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Dayend.DailyUpdateincBackup(pDate)", pDate
End Sub
