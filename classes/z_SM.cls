VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_SM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub CancelRedeemedDeposit(pEXCHNum As Long)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CancelRedeemedDeposit"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pEXCHNUM", adInteger)
    cmd.Parameters.Append par
    par.Value = pEXCHNum
    
    cmd.ActiveConnection = oPC.DBLocalConn
    cmd.Execute
    
    Set cmd = Nothing

End Sub
Public Function FindPIDFromCode(pCODE As String) As String
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_FindPIDFromCode"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Code", adVarChar, adParamInput, 20, pCODE)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@PID", adGUID, adParamOutput)
    cmd.Parameters.Append par
    
    oPC.OpenLocalDatabase
    cmd.ActiveConnection = oPC.DBLocalConn
    cmd.Execute
    FindPIDFromCode = FNS(cmd.Parameters(1))
    Set cmd = Nothing

End Function


Public Sub MarkRedeemedDeposit(pCOLID As Long)
    oPC.DBLocalConn.Execute "UPDATE tCOL SET COL_DEPOSITSTATUS = 'E' WHERE COL_COLID = " & pCOLID
End Sub

Public Function CanRedeemDeposit(pCOLID As Long, pReplacedExchNum As Long) As Boolean
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CanRedeemDeposit"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@COLID", adInteger)
    cmd.Parameters.Append par
    par.Value = pCOLID
    
    Set par = cmd.CreateParameter("@EXCHNUM", adInteger)
    cmd.Parameters.Append par
    par.Value = pReplacedExchNum
    
    Set par = cmd.CreateParameter("@Ret", adBoolean, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.DBLocalConn
    cmd.Execute
    CanRedeemDeposit = FNB(cmd.Parameters(2).Value)
    
    Set cmd = Nothing
End Function

Public Function PayDeposit(pCOLID As Long)
    oPC.DBLocalConn.Execute "UPDATE tCOL SET COL_DEPOSITSTATUS = 'P' WHERE COL_COLID = " & pCOLID
End Function
Public Function RefundDeposit(pCOLID As Long)
    oPC.DBLocalConn.Execute "UPDATE tCOL SET COL_DEPOSITSTATUS = 'X' WHERE COL_COLID = " & pCOLID
End Function
Public Function DeleteBridge(INVID As Long)
    oPC.DBLocalConn.Execute "DELETE FROM tBRIDGE WHERE INV_ID = " & INVID
End Function
Public Function GetStationNames() As ADODB.Recordset
    Dim lngResult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset
    Set cmd = New ADODB.Command
    cmd.CommandText = "GetTillStationNames"
    cmd.CommandType = adCmdStoredProc
    
    cmd.ActiveConnection = oPC.DBMainConn
    Set rs = cmd.Execute
    Set cmd = Nothing
    Set par = Nothing
    Set GetStationNames = rs
    Set rs = Nothing

End Function
'GetExchanges
Public Function GetExchanges(pStationName As String, pExchangeNumber As String) As ADODB.Recordset
    Dim lngResult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset
    Set cmd = New ADODB.Command
    cmd.CommandText = "GetExchanges"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@StationName", adVarChar, adParamInput, 40, Trim(pStationName))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ExchangeNumber", adInteger, adParamInput, , pExchangeNumber)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.DBMainConn
    Set rs = cmd.Execute
    Set cmd = Nothing
    Set par = Nothing
    Set GetExchanges = rs
    Set rs = Nothing

End Function
Public Function GetQuotationlines(pQuotationCode As String) As ADODB.Recordset
    Dim lngResult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset

    Set cmd = New ADODB.Command
    cmd.CommandText = "GetQuotationLines"
    cmd.CommandType = adCmdStoredProc
    
    
    Set par = cmd.CreateParameter("@TRCode", adVarChar, adParamInput, 20, pQuotationCode)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.DBMainConn
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.Open cmd, , adOpenDynamic, adLockOptimistic
    'Set rs = cmd.Execute
    Set cmd = Nothing
    Set par = Nothing
    Set GetQuotationlines = rs
   ' Set rs = Nothing

End Function

Public Function UpdateExchAck(pEXCHID As String)
    On Error Resume Next
    oPC.DBLocalConn.Execute "UPDATE tEXCHANGE SET EXCH_RECEIVED = 1 WHERE EXCH_ID = '" & pEXCHID & "'"
End Function
