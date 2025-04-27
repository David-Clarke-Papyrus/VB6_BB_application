VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Debtors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim tlDebtors As z_TextList

Public Sub AlDebtors(rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
Dim i As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "SELECT * FROM tTP WHERE TP_ROLE = 3"
    cmd.CommandType = adCmdText
    cmd.ActiveConnection = oPC.COShort
    
    
'    Set par = Nothing
'    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
'    cmd.Parameters.Append par
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic, adLockOptimistic

    Set cmd = Nothing
    Set rs.ActiveConnection = Nothing
'---------------------------------------------------
    'If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

'errHandler:
'    ErrorIn "z_Debtors.AlDebtors(rs)", rs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Debtors.AlDebtors(rs)", rs, EA_NORERAISE
End Sub
