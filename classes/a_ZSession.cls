VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_ZSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim dteFromDate As Date
Dim dteEndDate As Date
Dim strZSessionID As Variant
Dim oOpSession As z_OpSession
Dim dteNominal As Date
Dim oCol As Collection
Dim strSupervisorName As String
Dim lngSupervisorID As Long
Dim strOutBox As String

Public Sub Load(PID As String)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset

    oPC.OpenLocalDatabase
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * fROM tZSession LEFT JOIN tStaffMembers ON Z_SUPERVISORID = SM_ID WHERE z_ID = '" & PID & "'", oPC.DBLocalConn, adOpenKeyset, adLockOptimistic
    
    strSupervisorName = FNS(rs!SM_Name)
    lngSupervisorID = FNN(rs!Z_SUPERVISORID)
    dteNominal = FND(rs!Z_NOMINALDATE)
    dteFromDate = FND(rs!Z_STARTDATE)
    dteEndDate = FND(rs!Z_ENDDATE)
    strZSessionID = PID
 '   lngSequenceNumber = FNN(rs!Z_SequenceNumber)
    rs.Close
    Set rs = Nothing
    
    oPC.CloseLocalDatabase
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ZSession.Load(pID)", PID
End Sub
Public Property Get SupervisorID() As Long
    SupervisorID = lngSupervisorID
End Property
Public Property Let SupervisorID(val As Long)
    lngSupervisorID = val
    
    oPC.OpenLocalDatabase
    
    oPC.DBLocalConn.Execute "UPDATE tZSession SET Z_SUPERVISORID = " & lngSupervisorID
    
    oPC.CloseLocalDatabase
    
End Property
Public Property Let ClientOutbox(pOutbox As String)
    strOutBox = pOutbox
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
Dim str As String
    Set oOpSession = Nothing
    Set oOpSession = New z_OpSession
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Class_Initialize"
End Sub
Public Property Get SupervisorName() As String
    SupervisorName = strSupervisorName
End Property
Public Property Let SupervisorName(val As String)
    strSupervisorName = Trim(val)
End Property
Private Sub Class_Terminate()
    On Error GoTo errHandler
    Set oCol = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Class_Terminate"
End Sub

'Public Property Let ZT_DateTime(val As Date)
'    dtZT_DateTime = val
'End Property
'Public Property Get ZT_DateTime() As Date
'    ZT_DateTime = dtZT_DateTime
'End Property

Public Property Get FromDate() As Date
    FromDate = dteFromDate
End Property
Public Property Let FromDate(val As Date)
    dteFromDate = val
End Property
Public Property Get EndDate() As Date
    EndDate = dteEndDate
End Property
Public Property Let EndDate(val As Date)
    dteEndDate = val
End Property


Public Function AddZExchange() As d_ZExchange
    On Error GoTo errHandler
Dim oZExch As New d_ZExchange
Dim oGUID As New CGuid

    oZExch.ExchangeGUID = oGUID.GetGuid
    oZExch.Index = oCol.Count + 1
    oCol.Add Item:=oZExch, Key:=oZExch.Index
    Set AddZExchange = oZExch
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.AddZExchange"
End Function

Public Function ZExchList() As Collection
    On Error GoTo errHandler
    Set ZExchList = oCol
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.ZExchList"
End Function

Public Function LoadPrevZActions(ByVal FromDate As Date) As Integer
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.LoadPrevZActions(FromDate)", FromDate
End Function
Public Property Get NominalDate() As Date
    NominalDate = dteNominal
End Property
Public Property Get NominalDateF() As String
    NominalDateF = Format(dteNominal, "dd/mm/yyyy")
End Property
Public Function Start_Z_Session(pSMID As Long, pName As String) As Variant
On Error GoTo errHandler
Dim str As String
Dim oGUID As New CGuid
Dim rs As ADODB.Recordset
Dim tmpDate As Date
Dim strPos As String

    oPC.OpenLocalDatabase
strPos = "0"
    strZSessionID = oGUID.GetGuid
strPos = "1"
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tZSession WHERE Z_ID = null", oPC.DBLocalConn, adOpenDynamic, adLockOptimistic
strPos = "2"
    With rs
        .AddNew
strPos = "3"
        !Z_STARTDATE = Now()
        !Z_ID = strZSessionID
        !Z_SUPERVISORID = pSMID
        !Z_STATUS = "S"
        dteNominal = Date
        dteFromDate = FND(!Z_STARTDATE)
        !Z_NOMINALDATE = dteNominal
        !Z_TILLPOINT = oPC.StationName
        .Update
strPos = "4"
    End With
    rs.Close
strPos = "5"
    Set rs = Nothing
    strSupervisorName = pName
    dteNominal = Date
strPos = "6"
    str = "Update tAppSettings Set OpenZSessionID = '" & strZSessionID & "'"
    oPC.DBLocalConn.Execute str
strPos = "7"
    Start_Z_Session = strZSessionID
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Start_Z_Sezzion", , , , "strPOS", Array(strPos)
End Function

Public Function Close_Z_Session()
On Error GoTo errHandler
Dim str As String
Dim oGUID As New CGuid
Dim rs As ADODB.Recordset
Dim q As New MSMQQueueInfo
Dim RequestQ As MSMQQueue
Dim msg As New MSMQMessage
Dim strPos As String
 '''''
 'Dim ozconn As New z_POSCLIConnection
 '   MsgBox ozconn.ExchangeNumber
    oPC.OpenLocalDatabase
    
    If strZSessionID > "" Then
        str = "UPDATE tZSession Set Z_ENDDATE = '" & ReverseDateTime(Now()) & "', Z_STATUS = 'X' WHERE Z_ID = '" & strZSessionID & "'"
        oPC.DBLocalConn.Execute str

        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT 'Z' as typ,Z_ID,Z_STARTDATE,Z_ENDDATE,Z_TillPoint FROM tZSession WHERE Z_ID = '" & strZSessionID & "'", oPC.DBLocalConn, adOpenStatic
        
        OpSession.Close_OP_Session True
        If Not rs.EOF Then
          '  q.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\QPOS"
            If IsNumeric(Replace(oPC.ServerIPAddress, ".", "")) Then
                q.FormatName = "DIRECT=TCP:" & oPC.ServerIPAddress & "\Private$\QPOS"
            Else
                q.FormatName = "DIRECT=OS:" & oPC.ServerIPAddress & "\Private$\QPOS"
            End If
            Set RequestQ = q.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            msg.Delivery = MQMSG_DELIVERY_RECOVERABLE
            msg.MaxTimeToReachQueue = oPC.POSMessageTimeout
            msg.Label = "ZEND," & Format(Now, "dd/mm/yyyy HH:NN")
            msg.Body = "Z" & vbTab & FNS(rs!Z_ID) & vbTab & FNS(rs!Z_TILLPOINT) & vbTab & ReverseDateTime(rs!Z_STARTDATE) & vbTab & ReverseDateTime(rs!Z_ENDDATE) & vbTab & CStr(lngExchangeNumber)
            msg.Send RequestQ
        End If
        rs.Close
        Set rs = Nothing
        strZSessionID = ""
        str = "Update tAppSettings Set OpenZSessionID = NULL"
        oPC.DBLocalConn.Execute str
    End If
    
    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Close_Z_Session", , , , "strPOS", Array(strPos)
End Function

Public Property Get Current_Z_Session_ID() As Variant
    Current_Z_Session_ID = strZSessionID
End Property

Public Function InSession() As Boolean
    InSession = False
    If Not IsNull(strZSessionID) Then
        If strZSessionID > "" Then
            InSession = True
        End If
    End If
End Function

Public Property Get OpSession() As z_OpSession
    Set OpSession = oOpSession
End Property

Public Function LoadOpenXSession() As Boolean
    On Error GoTo errHandler
Dim rs As ADODB.Recordset

    oPC.OpenLocalDatabase
    
    LoadOpenXSession = False
    Set rs = New ADODB.Recordset
    rs.Open "SELECT SM_Name,tOpSession.* fROM tOpSession LEFT JOIN tStaffMembers ON OPS_OperatorID = SM_ID WHERE ISNULL(OPS_ENDTIME,0) < '2000-01-01' AND OPS_Z_ID = '" & strZSessionID & "'", oPC.DBLocalConn, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then
        Set oOpSession = New z_OpSession
        oOpSession.Current_OP_Session_ID = FNS(rs!OPS_ID)
        oOpSession.OperatorID = FNN(rs!OPS_OPERATORID)
        oOpSession.Name = FNS(rs!SM_Name)
        oOpSession.DateStarted = FND(rs!OPS_STARTTIME)
        oOpSession.DateEnded = FND(rs!OPS_endtime)
        oOpSession.FloatValue = FNDBL(rs!OPS_FloatValue)
        oOpSession.FloatBreakdown = FNS(rs!OPS_FloatBreakdown)
        LoadOpenXSession = True
    End If
    rs.Close
    Set rs = Nothing
    
    oPC.CloseLocalDatabase
    
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ZSession.LoadOpenXSession"
End Function
