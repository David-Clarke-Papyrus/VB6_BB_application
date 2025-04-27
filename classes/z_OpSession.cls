VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_OpSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mOperatorID As Long
Dim mSupervisorID As Long
Dim strOPSessionID As Variant
Dim strName As String
Dim mDateStarted As Date
Dim mDateEnded As Date
Dim mFloat As Double
Dim mFloatBreakdown As String


Public Property Get OPSID() As String
    OPSID = strOPSessionID
End Property
Public Property Let Name(val As String)
    strName = Trim(val)
End Property
Public Property Get Name() As String
    Name = Trim(strName)
End Property
Public Property Let DateStarted(val As Date)
Dim str As String
    
    mDateStarted = val
'   str = "UPDATE tOpSession Set OPS_StartTime = '" & ReverseDateTime(val) & "' WHERE OPS_ID = '" & Me.Current_OP_Session_ID & "'"
 '   oPC.DBLocalConn.Execute str
    
  '  oPC.CloseLocalDatabase
    
End Property
Public Property Get DateStarted() As Date
    DateStarted = mDateStarted
End Property
Public Property Let DateEnded(val As Date)
Dim str As String
    mDateEnded = val
'   oPC.OpenLocalDatabase

    str = "UPDATE tOpSession Set OPS_EndTime = '" & ReverseDateTime(val) & "' WHERE OPS_ID = '" & Me.Current_OP_Session_ID & "'"
    oPC.DBLocalConn.Execute str

  '  oPC.CloseLocalDatabase

End Property
Public Property Get DateEnded() As Date
    DateEnded = mDateEnded
End Property

Public Property Let OperatorID(val As Long)
Dim str As String
    mOperatorID = val
    str = "UPDATE tOpSession Set OPS_OperatorID = " & val & " WHERE OPS_ID = '" & Me.Current_OP_Session_ID & "'"
    
    oPC.DBLocalConn.Execute str
    
    
End Property
Public Sub SetOperatorID(val As Long, Optional FloatValue As Double, Optional sFloatBreakdown As String)
Dim str As String
    mOperatorID = val
    mFloat = FloatValue
    str = "UPDATE tOpSession Set OPS_OperatorID = " & CStr(val) & ", OPS_FloatValue = " & CStr(FloatValue) & ", OPS_Floatbreakdown= '" & sFloatBreakdown & "' WHERE OPS_ID = '" & Me.Current_OP_Session_ID & "'"
    
    oPC.DBLocalConn.Execute str

End Sub
Public Property Get FloatBreakdown() As String
    FloatBreakdown = mFloatBreakdown
End Property
Public Property Let FloatBreakdown(val As String)
    mFloatBreakdown = val
End Property

Public Property Get FloatValue() As Double
    FloatValue = mFloat
End Property
Public Property Let FloatValue(val As Double)
    mFloat = val
End Property

Public Property Get OperatorID() As Long
    OperatorID = mOperatorID
End Property
Public Property Get SupervisorID() As Long
    SupervisorID = mSupervisorID
End Property

Public Property Get Current_OP_Session_ID() As Variant
    Current_OP_Session_ID = strOPSessionID
End Property
Public Property Let Current_OP_Session_ID(val As Variant)
    strOPSessionID = val
End Property
Public Function Start_OP_Session_OLD(pZSessionID As String, pOPID As Long)
On Error GoTo errHandler
Dim str As String
Dim rsOP As ADODB.Recordset
Dim oGUID As New CGuid

    oPC.OpenLocalDatabase
    
    strOPSessionID = oGUID.GetGuid
    Set rsOP = New ADODB.Recordset
    rsOP.Open "SELECT * FROM tOPSession WHERE OPS_ID = NULL", oPC.DBLocalConn, adOpenKeyset, adLockOptimistic
    With rsOP
        .AddNew
        !OPS_ID = strOPSessionID
        mDateStarted = Now()
        DateStarted = mDateStarted
        !OPS_STARTTIME = DateStarted
        !OPS_Z_ID = pZSessionID
        !OPS_STATUS = "S"
        !OPS_OPERATORID = pOPID
        !OPS_FloatValue = mFloat
        !OPS_FloatBreakdown = mFloatBreakdown
        .Update
    End With
 '   rsOP.MoveLast
 '   strOPSessionID = rsOP!OPS_ID
    str = "Update tAppSettings Set OpenOpSessionID = '" & strOPSessionID & "'"
    oPC.DBLocalConn.Execute str
    Set rsOP = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Start_OP_Sezzion"
End Function
Public Function Start_OP_Session(pZSessionID As String, pOPID As Long)
On Error GoTo errHandler
Dim str As String
Dim rsOP As ADODB.Recordset
Dim oGUID As New CGuid

    oPC.OpenLocalDatabase
    
    strOPSessionID = oGUID.GetGuid
    DateStarted = Now()

    str = "Insert into tOPSession (OPS_ID,OPS_STARTTIME,OPS_Z_ID,OPS_STATUS,OPS_OPERATORID,OPS_FloatValue,OPS_FloatBreakdown) values " _
     & "( '" & strOPSessionID & "'," _
     & "'" & ReverseDateTime(DateStarted) & "'," _
     & "'" & pZSessionID & "'," _
     & "'" & "S" & "'," _
     & pOPID & "," _
     & mFloat _
     & ",'" & "0,0,0,0,0,0,0,0,0,0,0,0" & "'" _
     & ")"
    oPC.DBLocalConn.Execute str
    str = "Update tAppSettings Set OpenOpSessionID = '" & strOPSessionID & "'"
    oPC.DBLocalConn.Execute str
    Set rsOP = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Start_OP_Sezzion"
End Function


Public Function Close_OP_Session(bFromZClose As Boolean)
On Error GoTo errHandler
Dim str As String
Dim rs As ADODB.Recordset
Dim q As New MSMQQueueInfo
Dim RequestQ As MSMQQueue
Dim msg As New MSMQMessage

    
    If strOPSessionID > "" Then
        str = "Update tOPSession Set OPS_EndTime = '" & ReverseDateTime(Now) & "',OPS_STATUS = 'X' WHERE OPS_ID = '" & strOPSessionID & "'"
        If Not bFromZClose Then
            oPC.OpenLocalDatabase
        End If
        oPC.DBLocalConn.Execute str
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT 'O' as typ,OPS_ID,OPS_ENDTIME,OPS_FLOATVALUE,OPS_FLOATBREAKDOWN FROM tOPSession WHERE OPS_ID = '" & strOPSessionID & "'", oPC.DBLocalConn, adOpenStatic
        
        Set rs.ActiveConnection = Nothing
        
        If Not rs.EOF Then
         '   q.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\QPOS"
            If IsNumeric(Replace(oPC.ServerIPAddress, ".", "")) Then
                q.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\QPOS"
            Else
                q.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\QPOS"
            End If
            
            
            Set RequestQ = q.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            msg.Delivery = MQMSG_DELIVERY_RECOVERABLE
            msg.MaxTimeToReachQueue = oPC.POSMessageTimeout
            msg.Label = "XEND," & Format(Now, "dd/mm/yyyy HH:NN")
            msg.Body = "O" & vbTab & "" & vbTab & "" & vbTab & CDate(0) & vbTab & CDate(0) & vbTab & ReverseDateTime(CDate(0)) & vbTab & FNS(rs!OPS_ID) & vbTab & ReverseDateTime(CDate(0)) & vbTab & ReverseDateTime(rs!OPS_endtime) & vbTab & (rs!OPS_FloatValue) & vbTab & (rs!OPS_FloatBreakdown)
            msg.Send RequestQ
        End If
        
        
        str = "Update tAppSettings Set OpenOPSessionID = NULL"
        oPC.DBLocalConn.Execute str
        strOPSessionID = ""
    End If
    
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Close_Z_Session"
End Function

Public Function InSession() As Boolean
    InSession = False
    If Not IsNull(strOPSessionID) Then
        If strOPSessionID > "" Then
            InSession = True
        End If
    End If
End Function

Public Sub Load(pOpSessionID As String)

End Sub

