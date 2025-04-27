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

Public DBConn As ADODB.Connection



Public Function dbConnect(ByVal pUsername As String, ByVal pPassword As String, DBPath As String, _
            DBName As String) As Integer
On Error GoTo ERR_Handler

Dim szConnect As String
    dbConnect = 0
    If DBConn Is Nothing Then
        szConnect = "Provider=Microsoft.jet.OLEDB.4.0;DATA SOURCE=" & DBPath & DBName _
                            & ";JET OLEDB:SYSTEM DATABASE=" & DBPath & "System.mdw" & ";USER ID = " _
                            & pUsername & ";Password=" & pPassword & ";"
'        szConnect = "Provider=MSDASQL;DSN=PapyPos;UID=" & pUsername & ";pwd=" & pPassword & ";"
        Set DBConn = New ADODB.Connection

        DBConn.Open szConnect
    End If
EXIT_Handler:
    Exit Function
ERR_Handler:
    If DBConn.Errors.Count > 0 Then
    dbConnect = DBConn.Errors.Item(0).NativeError
        GoTo EXIT_Handler
    Else
        App.LogEvent "dbConnect - " & Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
        Err.Raise vbObjectError
    End If
    

End Function
Public Function dbCloseConnect() As Integer
On Error GoTo ERR_Handler

    DBConn.Close
    Set DBConn = Nothing
EXIT_Handler:
    Exit Function

ERR_Handler:
    If DBConn.Errors.Count > 0 Then
    dbCloseConnect = DBConn.Errors.Item(0).NativeError
        GoTo EXIT_Handler
    Else
        App.LogEvent "dbCloseConnect - " & Err.Description & " (" & Err.Number & ")", vbLogEventTypeError
        Err.Raise vbObjectError
    End If
    
End Function


