VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_DatabasePersist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dbConn As ADODB.Connection

Dim sDBName As String
Private Sub Class_Initialize()
    sDBName = "NEWDB"
End Sub

Public Property Get DBName() As String
    DBName = sDBName
End Property

Public Function dbConnect(ByVal pUsername As String, ByVal pPassword As String, Optional sProvider As String) As Integer
On Error GoTo ERR_Handler

Dim szConnect As String
    dbConnect = 0
    If dbConn Is Nothing Then
        If sProvider <> "" Then
            szConnect = "Provider=" & sProvider & ";DSN=" & sDBName & ";UID=" & pUsername & ";pwd=" & pPassword & ";"
            Set dbConn = New ADODB.Connection

            dbConn.Open szConnect
        Else
            Set dbConn = New ADODB.Connection
            dbConn.Provider = "sqloledb"
         '   szConnect = "Provider=sqloledb;DSN=" & sDBName & ";UID=" & pUsername & ";pwd=" & pPassword & ";"
            dbConn.Open "Database=NEWDB;User Id=sa;Password=CITROEN;"
        End If
   '     Set dbConn = New ADODB.Connection

        'dbConn.Open szConnect

    End If
EXIT_Handler:
    Exit Function
ERR_Handler:
    If dbConn.Errors.Count > 0 Then
    dbConnect = dbConn.Errors.Item(0).NativeError
        GoTo EXIT_Handler
    Else
        App.LogEvent "dbConnect - " & Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
        Err.Raise vbObjectError
    End If
    

End Function
Public Function dbCloseConnect() As Integer
On Error GoTo ERR_Handler

    dbConn.Close
    Set dbConn = Nothing

EXIT_Handler:
    Exit Function

ERR_Handler:
    If dbConn.Errors.Count > 0 Then
    dbCloseConnect = dbConn.Errors.Item(0).NativeError
        GoTo EXIT_Handler
    Else
        App.LogEvent "dbCloseConnect - " & Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
        Err.Raise vbObjectError
    End If
    
End Function

Public Property Get IsConnected() As Boolean
    If Not dbConn Is Nothing Then
        IsConnected = (dbConn.State = adStateOpen)
    End If
End Property

