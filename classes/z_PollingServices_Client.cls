VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_PollingServices_Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Dim WithEvents oPoll As clsPoll
Event ConnectionStatus(iStatus As Integer)
Event UpdatingLocalDatabase(bOn As Boolean, lngCnt As Long)
Event Counter(lngCnt As Long)
Event Ack()
Dim iExchange As Integer
Dim iSaleLine As Integer
Dim iPayment As Integer
Dim bRSOK As Boolean
Dim sGUID As String
Dim oGD As z_GetData

Dim sPOSSQLServer As String

Dim bPollingActive As Boolean

Private QI As MSMQQueueInfo


Private QPOSACK As MSMQQueue



Private Sub Class_Initialize()

    On Error GoTo errHandler
    
    Set oPS = Me
    
    Set QI = New MSMQQueueInfo
    QI.PathName = oPC.NameOfPC & "\Private$\QPOSACK"
    Set QPOSACK = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set POSACKEvent = New MSMQEvent
    QPOSACK.EnableNotification POSACKEvent

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.Class_Initialize"
End Sub




Public Property Get guid() As String
    On Error GoTo errHandler
    guid = sGUID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "clsExchange.guid"
End Property



Private Sub oPoll_ConnectionStatus(i As Integer)
    RaiseEvent ConnectionStatus(i)
End Sub



