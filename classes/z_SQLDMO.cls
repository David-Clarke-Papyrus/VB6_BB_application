VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_SQLDMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Dim oDatabase As SQLDMO.Database2
'Dim oSQLServer As SQLDMO.SQLServer2


Public Sub RebuildIndexes()
'Dim oTable As SQLDMO.Table
'    For Each oTable In oDatabase.Tables
'        If Not oTable.SystemObject Then oTable.RebuildIndexes
'    Next
End Sub
Public Sub ShrinkDatabase()
'    oDatabase.Shrink 10, SQLDMOShrink_Default
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
'    Set oSQLServer = New SQLDMO.SQLServer
'    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
'    With oSQLServer
'        .LoginSecure = False
'        .AutoReConnect = False
'        .Connect oPC.servername, "sa", oPC.Password
'    End With
'
'    Set oDatabase = oSQLServer.Databases("PBKS")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQLDMO.ClassInitialize", , , , "Servername,Password", Array(oPC.servername, oPC.Password)

End Sub


'================
Public Sub SwitchPOSTriggers(Switch As Boolean)
    On Error GoTo errHandler
'    Set oSQLServer = New SQLDMO.SQLServer
'    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
'    With oSQLServer
'        .LoginSecure = False
'        .AutoReConnect = False
'        .Connect oPC.servername, "sa", oPC.Password
'    End With
'    Set oDatabase = oSQLServer.Databases("PBKS")
'    oDatabase.Tables("tAPPL").Triggers("trigAPPUpdate").Enabled = Switch
'    oDatabase.Tables("tMarketing").Triggers("trigMarketingDelete").Enabled = Switch
'    oDatabase.Tables("tMarketing").Triggers("trigMarketingUpdate").Enabled = Switch
'    oDatabase.Tables("tProductSection").Triggers("trigSectChange").Enabled = Switch
'    oDatabase.Tables("tStaffmember").Triggers("trigStaffmemberUpdate").Enabled = Switch
'    oDatabase.Tables("tTP").Triggers("trigCustomerDelete").Enabled = Switch
'    oDatabase.Tables("tTP").Triggers("trigCustomerUpdate").Enabled = Switch
'    Set oDatabase = Nothing
'    oSQLServer.Close
'    Set oSQLServer = Nothing
    Exit Sub
errHandler:
    ErrPreserve
    If Err = -2147199728 Then  'If trigger is missing
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "PapyConn.SwitchPOSTriggers(Switch)", Switch
End Sub
Public Sub RestoreDatabase(pFilename As String, Optional pName As String = "PBKS_TEST")
'Dim oRestore As New SQLDMO.Restore
'    On Error Resume Next
'    oSQLServer.DetachDB pName
'    On Error GoTo errHandler
'    With oRestore
'    'this is where your backup will be restored to
'    .Database = pName
'    'same as EM or TSQL, you can restore database, file, or log, here we're going to
'    'use database
'    .Action = SQLDMORestore_Database
'    'this is the "force restore over existing database" option
'    .ReplaceDatabase = True
'    'this does a restore from a file instead of a device - note that we're still
'    'restoring a database, NOT a file group
'    .RelocateFiles = "[PBKS_Data]" + "," + "[C:\PBKS\DATA\PBKSTEST_DATA.mdf]" _
'        + "," + "[PBKS_LOG]" + "," + "[C:\PBKS\DATA\PBKSTEST_LOG.ldf]"
'
'    .Files = pFilename
'    'do it
'    .SQLRestore oSQLServer
'    End With
'
'    'standard clean up
'    Set oRestore = Nothing
'    oSQLServer.Disconnect
'    Set oSQLServer = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQLDMO.RestoreDatabase(pFilename)", pFilename
End Sub

Public Sub RestoreFDDatabase(pFilename As String, Optional pName As String = "PBKSFD")
'Dim oRestore As New SQLDMO.Restore
'    On Error Resume Next
'    oSQLServer.DetachDB pName
'    On Error GoTo errHandler
'    With oRestore
'    'this is where your backup will be restored to
'    .Database = pName
'    'same as EM or TSQL, you can restore database, file, or log, here we're going to
'    'use database
'    .Action = SQLDMORestore_Database
'    'this is the "force restore over existing database" option
'    .ReplaceDatabase = True
'    'this does a restore from a file instead of a device - note that we're still
'    'restoring a database, NOT a file group
'    .RelocateFiles = "[PBKSFD_Data]" + "," + "[C:\PBKS\DATA\PBKSFD_DATA.mdf]" _
'        + "," + "[PBKSFD_LOG]" + "," + "[C:\PBKS\DATA\PBKSFD_LOG.ldf]"
'
'    .Files = pFilename
'    'do it
'    .SQLRestore oSQLServer
'    End With
'
'    'standard clean up
'    Set oRestore = Nothing
'    oSQLServer.Disconnect
'    Set oSQLServer = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQLDMO.RestoreFDDatabase(pFilename)", pFilename
End Sub

