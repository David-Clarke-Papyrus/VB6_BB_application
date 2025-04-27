VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_POSCLIConnectionShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public DBConn As ADODB.Connection

Public Function dbConnecttoShape()
Dim strConnection As String

    On Error GoTo errHandler

    If DBConn Is Nothing Then
        Set DBConn = New ADODB.Connection
    End If
    If DBConn.State <> 0 Then DBConn.Close
   ' DBConn.Provider = "MSDataShape"
    DBConn.ConnectionTimeout = 120
    strConnection = "Provider=MSDataShape;Data provider=sqloledb;Data Source=" & oPC.LocalSQLServerName & ";Initial Catalog=PBKSFD;User Id=sa" & ";Password=" & oPC.Password & ";"""
    DBConn.Open strConnection
    DBConn.CommandTimeout = 120
    dbConnecttoShape = 0
    
EXIT_Handler:
    Exit Function
errHandler:
    ErrPreserve
    dbConnecttoShape = 1
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSCLIConnection.dbConnecttoShape", , , , "strConnection", Array(strConnection)
End Function

Public Function dbCloseConnectShape() As Integer
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

