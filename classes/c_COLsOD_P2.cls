VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_COLsOD_P2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event lngProgress(p As Long)
Event lngMax(p As Long)
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset


Public Function Fetch(COLS As ADODB.Recordset, POLS As ADODB.Recordset, COLActs As ADODB.Recordset, Optional pETABefore As Date, _
        Optional pTPID As Long, Optional pPID As String, Optional pSTAFFID As Long, Optional pStatusChangeDate As Date) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim OpenResult As Integer

Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    If pPID > "" Then
        cmd.CommandText = "vGetCOLSOS"
    Else
        cmd.CommandText = "vGetCOLSOS2"
    End If
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput)
    prm.Value = pTPID
    cmd.Parameters.Append prm
    
    Set prm = cmd.CreateParameter("@STAFFID", adInteger, adParamInput)
    prm.Value = pSTAFFID
    cmd.Parameters.Append prm
    
    If pETABefore <= CDate(1990 - 1 - 1) Then
        Set prm = cmd.CreateParameter("@PRIORTO", adDate, adParamInput, , ReverseDate(CDate("2200-01-01")))
    Else
        Set prm = cmd.CreateParameter("@PRIORTO", adDate, adParamInput, , ReverseDate(pETABefore))
    End If
    cmd.Parameters.Append prm
    
    If pPID > "" Then
        Set prm = cmd.CreateParameter("@PID", adVarChar, adParamInput, 40)
        prm.Value = pPID
        cmd.Parameters.Append prm
    End If
    
    Set prm = cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
    cmd.Parameters.Append prm
    
    Set prm = cmd.CreateParameter("@STATUSCHANGEDATE", adDate, adParamInput)
    If pStatusChangeDate > CDate(0) Then
        prm.Value = pStatusChangeDate
    End If
    cmd.Parameters.Append prm
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic, adLockOptimistic
    Set rs.ActiveConnection = Nothing
    Set COLS = rs
'================
    If pPID > "" Then
        cmd.CommandText = "vGetCOLSOS_MatchingPOLS"
    Else
        cmd.CommandText = "vGetCOLSOS2_MatchingPOLS"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic, adLockOptimistic
    Set rs.ActiveConnection = Nothing
    Set POLS = rs
    
'================
    If pPID > "" Then
        cmd.CommandText = "vGetCOLSOS_COLActs"
    Else
        cmd.CommandText = "vGetCOLSOS2_COLActs"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.open cmd, , adOpenStatic, adLockOptimistic
    Set rs.ActiveConnection = Nothing
    Set COLActs = rs
    
    Set cmd = Nothing
    
EXIT_Handler:
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_COLsOD_P2.Fetch(pETABefore,pTPID,pPID,pStaffID)", Array(pETABefore, pTPID, pPID, _
         pSTAFFID)
End Function




