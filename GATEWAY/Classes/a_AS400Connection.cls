VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_AS400connection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private dteTimeStarted As Date
Private strPapyConnErr As String
Private strServer As String
Private strUsername As String
Private strPwd As String
Private strServerRootPath As String
Private strSharedFolderRoot As String


Public Function Connect(pConn As String) As Integer
    On Error GoTo errHandler
Dim strConn As String
    Connect = 99
    If cnAS400 Is Nothing Then Set cnAS400 = New ADODB.Connection
 '   strConn = "DSN=PT_1;UID=;PWD=;"
    cnAS400.ConnectionTimeout = 0

    cnAS400.Open pConn
    
EXIT_Handler:
    Connect = 0
'ERR_Handler:
'    MsgBox "Error in Connect: " & Error
'    oError.SetError "Diagnostic", CStr(Err) & Error, Now(), "DLL:Papyconn:OpenDB:in error handler", "", ""
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_AS400connection.Connect(pConn)", pConn, EA_NORERAISE
    
End Function

Public Sub Disconnect()
    On Error GoTo errHandler
    cnAS400.Close
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_AS400connection.Disconnect"
    
End Sub

Public Property Get Connection() As ADODB.Connection
    On Error GoTo errHandler
    Set Connection = cnAS400
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_AS400connection.Connection"
    
End Property
