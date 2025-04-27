VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_POSUtiltties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Connect() As Boolean
Dim iR As Integer
    If oPC Is Nothing Then
        Set oPC = New z_POSCLIConnection
        iR = oPC.dbConnect()
    End If
    Exit Function
End Function

Public Function DisConnect() As Boolean
    If Not oPC Is Nothing Then
        DisConnect = (oPC.dbCloseConnect = 0)
        Set oPC = Nothing
    End If
    Exit Function
End Function


Private Sub Class_Terminate()
    DisConnect
    Exit Sub
End Sub


Public Function ClearStaffMembers()
On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE * FROM tblStaffMembers"
    Connect
    oPC.DBConn.Execute strSQL
    DisConnect
    ClearStaffMembers = True
    Exit Function
errHandler:
    ErrPreserve
    ClearStaffMembers = False
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSUtilities.ClearStaffMembers"
End Function
Public Function ClearProducts()
On Error GoTo errHandler
Dim strSQL As String
    strSQL = "DELETE * FROM Product"
    Connect
    oPC.DBConn.Execute strSQL
    DisConnect
    ClearProducts = True
    Exit Function
errHandler:
    ErrPreserve
    ClearProducts = False
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSUtilities.ClearProducts"
End Function

Public Function ClearCustomers()
On Error GoTo errHandler
Dim strSQL As String
    
    strSQL = "DELETE * FROM tblCUstomer"
    Connect
    oPC.DBConn.Execute strSQL
    DisConnect
    ClearCustomers = True
    Exit Function
errHandler:
    ErrPreserve
    ClearCustomers = False
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_POSUtilities.ClearCustomers"
End Function

