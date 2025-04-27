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
Dim oDatabase As SQLDMO.Database2
Dim oSQLServer As SQLDMO.SQLServer2


Public Sub RebuildIndexes()
Dim oTable As SQLDMO.Table
    For Each oTable In oDatabase.Tables
        If Not oTable.SystemObject Then oTable.RebuildIndexes
    Next
End Sub
Public Sub ShrinkDatabase()
    oDatabase.Shrink 10, SQLDMOShrink_Default
End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set oSQLServer = New SQLDMO.SQLServer2
    oSQLServer.LoginTimeout = 0 '-1 is the ODBC default (60) seconds
    With oSQLServer
        .LoginSecure = False
        .AutoReConnect = False
        .Connect strSQLServerName, "sa", strPassword
    End With
    
  '  Set oDatabase = oSQLServer.Databases("PBKS")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQLDMO.ClassInitialize", , , , "Servername,Password", Array(strSQLServerName, strPassword)

End Sub


Public Sub RestoreDatabase(pFilename As String)
    On Error GoTo errHandler
Dim oRestore As New SQLDMO.Restore2
With oRestore
'this is where your backup will be restored to
        .Database = "PBKSTEST"
        'same as EM or TSQL, you can restore database, file, or log, here we're going to
        'use database
        .Action = SQLDMORestore_Database
        'this is the "force restore over existing database" option
        .ReplaceDatabase = True
        'this does a restore from a file instead of a device - note that we're still
        'restoring a database, NOT a file group
        .RelocateFiles = "[PBKS_Data]" + "," + "[C:\PBKS\DATA\PBKSTEST_DATA.mdf]" _
            + "," + "[PBKS_LOG]" + "," + "[C:\PBKS\DATA\PBKSTEST_LOG.ldf]"
        
        .Files = pFilename
        'do it
        .SQLRestore oSQLServer
End With
 
'standard clean up
    Set oRestore = Nothing
    oSQLServer.DisConnect
    Set oSQLServer = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_SQLDMO.RestoreDatabase(pFilename)", pFilename
End Sub


