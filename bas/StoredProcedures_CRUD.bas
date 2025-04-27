Attribute VB_Name = "StoredProcedures_CRUD"
Option Explicit

Public Function SaveApproReturn(IsNew As Boolean, TRID As Long, TPID As Long, DOCDate As Date, DOCCode As String, _
    Memo As String, Status As Integer, StaffID As Long) As Boolean
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "SaveApproReturn"
    cmd.CommandType = adCmdStoredProc

    Set par = cmd.CreateParameter("@IsNew", adBoolean, , , IsNew)
    cmd.Parameters.Append par

    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInputOutput, , TRID)
    cmd.Parameters.Append par

    Set par = cmd.CreateParameter("@TPID", adInteger, , , TPID)
    cmd.Parameters.Append par

    Set par = cmd.CreateParameter("@DocDate", adDate, , , DOCDate)
    cmd.Parameters.Append par

    Set par = cmd.CreateParameter("DocCode", adVarChar, , 20, DOCCode)
    cmd.Parameters.Append par

    Set par = cmd.CreateParameter("@Memo", adVarChar, , 500, Memo)
    cmd.Parameters.Append par

    Set par = cmd.CreateParameter("@Status", adSmallInt, , , Status)
    cmd.Parameters.Append par

    Set par = cmd.CreateParameter("@Memo", adInteger, , , StaffID)
    cmd.Parameters.Append par

    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    TRID = cmd.Parameters(1).Value

    Set cmd = Nothing
    Set par = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
End Function

Public Sub ConsolidateApproLines_Svr(TRID As Long, bMadeChange As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.CommandText = "ConsolidateApproLines"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("Return", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    
    cmd.execute
    iReturn = cmd.Parameters("Return")
    If iReturn > 1 Then Err.Raise EXC_SP_FAILED, "a_APP_P:ConsolidateLines", "SP failed"
    bMadeChange = (iReturn = 1)
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP_P.ConsolidateApproLines_Svr(TRID,bMadeChange)", Array(TRID, bMadeChange)
End Sub

