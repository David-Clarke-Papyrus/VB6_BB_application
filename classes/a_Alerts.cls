VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Alerts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function LoadAlert(MsgType As String, MsgText As String, Optional AcNo As String, Optional ProductID As String)
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim bRestart As Boolean

    bRestart = False
restart:
'-------------------------------
    OpenResult = oPC.OpenDBSHort(bRestart)
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "AlertLoad"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@MsgText", adVarChar, adParamInput, 500, MsgText)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@MsgType", adVarChar, adParamInput, 3, MsgType)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Acno", adVarChar, adParamInput, 15, AcNo)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'errHandler:
'    ErrorIn "z_SQL.SendAlert(MsgType,MsgText,DestTPID,ProductID)", Array(MsgType, MsgText, DestTPID, _
'         ProductID)
    Exit Function
errHandler:
    ErrorIn "z_SQL.LoadAlert(MsgType,MsgText,Acno,ProductID)", Array(MsgType, MsgText, AcNo, _
         ProductID)
End Function


