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
Attribute VB_Exposed = True
Option Explicit

Dim dtFromDate As Date
Dim sTillCode As String * 15
Dim strZSessionID As Variant
Dim oOpSession As z_OpSession
Dim dteNominal As Date
Dim oCol As Collection
Dim strSupervisorName As String
Dim lngSupervisorID As Long
Public Sub Load(pID As String)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * fROM tZSession LEFT JOIN tStaffMembers ON Z_SUPERVISORID = SM_ID WHERE z_ID = '" & pID & "'", oPC.DBConn, adOpenKeyset, adLockOptimistic
    strSupervisorName = FNS(rs!SM_Name)
    lngSupervisorID = FNN(rs!Z_SUPERVISORID)
    dteNominal = FND(rs!Z_NominalDate)
 '   lngSequenceNumber = FNN(rs!Z_SequenceNumber)
    rs.Close
    Set rs = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_ZSession.Load(pID)", pID
End Sub
Public Property Get SupervisorID() As Long
    SupervisorID = lngSupervisorID
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
    FromDate = dtFromDate
End Property
Public Property Let FromDate(val As Date)
    dtFromDate = val
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

    strZSessionID = oGUID.GetGuid
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tZSession WHERE Z_ID = null", oPC.DBConn, adOpenKeyset, adLockOptimistic
    With rs
        .AddNew
        !Z_STartDate = Now()
        !Z_ID = strZSessionID
        !Z_SUPERVISORID = pSMID
        !Z_STATUS = "S"
        dteNominal = Date
        !Z_NominalDate = dteNominal
        !Z_TillPoint = oPC.TillCode
        .Update
    End With
    rs.Close
    Set rs = Nothing
    strSupervisorName = pName
    dteNominal = Date
    str = "Update tAppSettings Set EX_Number = 1, OpenZSessionID = '" & strZSessionID & "'"
    oPC.DBConn.Execute str
    Start_Z_Session = strZSessionID
    
'    OpSession.Start_OP_Session (Me.Current_Z_Session_ID)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Start_Z_Sezzion"
End Function

Public Function Close_Z_Session()
On Error GoTo errHandler
Dim str As String
Dim oGUID As New CGuid
Dim rs As ADODB.Recordset
Dim sFileName As String

    If strZSessionID > "" Then
        str = "UPDATE tZSession Set Z_ENDDATE = '" & ReverseDateTime(Now()) & "', Z_STATUS = 'X' WHERE Z_ID = '" & strZSessionID & "'"
        oPC.DBConn.Execute str
        
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open "SELECT 'Z' as typ,Z_ID,Z_ENDDATE FROM tZSession WHERE Z_ID = '" & strZSessionID & "'", oPC.DBConn, adOpenStatic
        Set rs.ActiveConnection = Nothing
        If Not rs.EOF Then
            sFileName = oPC.TillCode & "-" & Format(Now(), "DDHHNNSS") 'Format(oGD.GetNextFileNum(), "00000")
            sFileName = "\" & sFileName & ".pos"
            rs.Save sRoot & sOutBox & sFileName, adPersistADTG
        End If
        
        OpSession.Close_OP_Session
        strZSessionID = ""
        str = "Update tAppSettings Set OpenZSessionID = NULL"
        oPC.DBConn.Execute str
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_SessionManager.Close_Z_Session"
End Function

Public Property Get Current_Z_Session_ID() As Variant
    Current_Z_Session_ID = strZSessionID
End Property
'Public Function CreateNewOpSessionID(pSMID As Long) As String
''Only create new session if last session is terminated
'
'    If OpSession.Current_OP_Session_ID > 0 Then
'        OpSession.Close_OP_Session
'        Set oOpSession = Nothing
'        Set oOpSession = New z_OpSession
'    End If
'    oOpSession.Start_OP_Session (Me.Current_Z_Session_ID)
'    OpSession.OperatorID = pSMID
'End Function


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
