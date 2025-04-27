VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_POLSOS2_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event lngProgress(p As Long)
Event lngMax(p As Long)
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset


Public Function Fetch(POLSOS As ADODB.Recordset, POLActions As ADODB.Recordset, Optional pETABefore As Date, Optional pTPID As Long, Optional pCOTPID As Long, Optional pPID As String, Optional pSTAFFID As Long, _
                        Optional pSubstitutions As Boolean, Optional pSkipactions As Boolean, Optional pStatusChangeDate As Date) As String
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
        cmd.CommandText = "vGetPOLSOS"
    Else
        cmd.CommandText = "vGetPOLSOS2"
    End If
    
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput)
    prm.Value = pTPID
    cmd.Parameters.Append prm
    
    Set prm = cmd.CreateParameter("@COTPID", adInteger, adParamInput, , pCOTPID)
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
    
    If pStatusChangeDate > CDate(0) Then
        Set prm = cmd.CreateParameter("@STATUSCHANGEDATE", adDate, adParamInput, , pStatusChangeDate)
        cmd.Parameters.Append prm
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockOptimistic
    Set rs.ActiveConnection = Nothing
    Set POLSOS = rs
'================
    If pPID > "" Then
        cmd.CommandText = "vGetPOLSOS_MatchingActions"
    Else
        cmd.CommandText = "vGetPOLSOS2_MatchingActions"
    End If
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenStatic, adLockOptimistic
    Set rs.ActiveConnection = Nothing
    Set POLActions = rs
     Set cmd = Nothing
EXIT_Handler:
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "c_POLSOS2_P.Fetch(POLSOS,POLActions,pETABefore,pTPID,pCOTPID,pPID,pStaffID," & _
'        "pSubstitutions,pSkipactions,POLSetXML)", Array(POLSOS, POLActions, pETABefore, pTPID, pCOTPID, pPID, _
'         pStaffID, pSubstitutions, pSkipactions, POLSetXML)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_POLSOS2_P.Fetch(POLSOS,POLActions,pETABefore,pTPID,pCOTPID,pPID,pStaffID," & _
        "pSubstitutions,pSkipactions)", Array(POLSOS, POLActions, pETABefore, pTPID, pCOTPID, pPID, pSTAFFID, _
         pSubstitutions, pSkipactions)
End Function


Public Function FetchPIDSet(POLSOS As ADODB.Recordset, POLActions As ADODB.Recordset, Optional POLSetXML As String) As String
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
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    If POLSetXML > "" Then
        cmd.CommandText = "vGetPOLSOS_perPO"
        Set prm = cmd.CreateParameter("@XMLArgs", adVarChar, adParamInput, -1, POLSetXML)
        cmd.Parameters.Append prm
    
        Set prm = cmd.CreateParameter("@COUNT", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open cmd, , adOpenStatic, adLockOptimistic
        Set rs.ActiveConnection = Nothing
        Set POLSOS = rs
        
        
        cmd.CommandText = "vGetPOLSOS_perPO_MatchingActions"
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open cmd, , adOpenStatic, adLockOptimistic
        Set rs.ActiveConnection = Nothing
        Set POLActions = rs
        Set cmd = Nothing
        
        
    End If
   
    
    
EXIT_Handler:
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "c_POLSOS2_P.Fetch(POLSOS,POLActions,pETABefore,pTPID,pCOTPID,pPID,pStaffID," & _
'        "pSubstitutions,pSkipactions,POLSetXML)", Array(POLSOS, POLActions, pETABefore, pTPID, pCOTPID, pPID, _
'         pStaffID, pSubstitutions, pSkipactions, POLSetXML)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "c_POLSOS2_P.FetchPIDSet(POLSOS,POLActions,POLSetXML)", Array(POLSOS, POLActions, _
         POLSetXML)
End Function







