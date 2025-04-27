VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Z_PollingServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'=================================================
'POSExchange variables from XML document
'ZSession
Dim strType As String
Dim strZID As String
Dim strTillpoint As String
Dim dteZStartDate As Date
Dim dteZEndDate As Date
Dim dteZNominalDate As Date
'OPSession
Dim dteOPStartDate As Date
Dim dteOPEndDate As Date
Dim strOPID As String
Dim lngOperatorID As Long
Dim lngSupervisorID As Long
'Exchange
Dim strEXCHID As String
Dim lngExchOperatorID As Long
Dim dteTimeOfSale As Date
Dim lngSaleValue As Long
Dim lngDiscountValue As Long
Dim lngVATValue As Long
Dim lngChangeGiven As Long
Dim lngLoyaltyValue As Long
Dim strExchType As String
Dim strNote As String
Dim lngVoids As Long
Dim lngTPID As Long
Dim lngExchangeNumber As Long

Dim TRID As Long
Dim lngErrstatus As Long
Dim lngErrPos As Long

'=================================================

Dim arQueuePath() As String

Public Event FileProcessed(Till As String) '0 = Other, 1 = Sale, 2 = Update
Public Event DataSent()
Public Event SwitchPolling(bOn As Boolean)
Dim bDBOK As Boolean

Private WithEvents POSEvent As MSMQEvent
Attribute POSEvent.VB_VarHelpID = -1
Private QPOS As MSMQQueue

'Private WithEvents REGEvent As MSMQEvent
'Private QREG As MSMQQueue
'
Private WithEvents SVREvent As MSMQEvent
Attribute SVREvent.VB_VarHelpID = -1
Private QSVR As MSMQQueue

Dim QI As MSMQQueueInfo
Dim QSVRAck As MSMQQueue
Dim SVRMsg As MSMQMessage

Event ErrorRegistering(pStationName As String)
Event ClientStatusChange()

Public Function PutMasterDataInOutbox() As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim cnt As Long
'-------------------------------
    OpenResult = oPC.OpenDBShort
'-------------------------------
      oPC.COShort.Execute "SELECT TOP 2000 tProdUpdates.* into #PRODS FROM tProdUpdates"
      oPC.COShort.Execute "SELECT PRU_EAN FROM #PRODs", cnt
      PlaceInQueue GetRecordset("SELECT * FROM #PRODS"), "PROD", "tProdUpdates"
      oPC.COShort.Execute "DROP table #PRODS"


      PlaceInQueue GetRecordset("SELECT * FROM tStaffMemberUpdate"), "STAF", "tStaffMemberUpdate"

      PlaceInQueue GetRecordset("SELECT  TOP 2000 * FROM tTPUpdate_Cust ORDER BY CU_Triggerdate"), "CUST", "tTPUpdate_Cust"

    '  PlaceInQueue GetRecordset("SELECT  TOP 2000 * FROM tCOUPDATE ORDER BY COU_TriggerDate"), "ORDR", "tCOUpdate"
      PlaceInQueue GetRecordset("SELECT  * FROM tCOUPDATE"), "ORDR", "tCOUpdate"

      PlaceInQueue GetRecordset("SELECT   * FROM tMarketing_Changes"), "MARK", "tMarketing_Changes"

      PlaceInQueue GetRecordset("SELECT * FROM tAPPUpdate"), "APPL", "tAPPUpdate"
    '  PlaceInQueue GetRecordset("SELECT  TOP 2000 * FROM tAPPUpdate ORDER BY APPL_TriggerDate"), "APPL", "tAPPUpdate"
          
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
        
    Exit Function
errHandler:

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.PutMasterDataInOutbox", , EA_ROLLBACK, oPC.COShort
End Function
Private Function GetRecordset(strSQL As String) As ADODB.Recordset
    On Error GoTo errHandler
Dim rs1 As ADODB.Recordset

        oPC.COShort.BeginTrans
        Set rs1 = New ADODB.Recordset
        rs1.CursorLocation = adUseClient
        rs1.Open strSQL, oPC.COShort, adOpenStatic, adLockOptimistic
        rs1.ActiveConnection = Nothing         'DISCONNECT RECORDSET
        Set GetRecordset = rs1
        Set rs1 = Nothing
        oPC.COShort.CommitTrans

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.GetRecordset(strSQL)", strSQL
End Function
Public Sub ClearCustomersOnFD()
    On Error GoTo errHandler
Dim j As Integer

        For j = 1 To UBound(arQueuePath)
            QI.FormatName = "DIRECT=OS:" & arQueuePath(j)
            Set QSVR = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            Set SVRMsg = New MSMQMessage
            SVRMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
            SVRMsg.Label = "ClearCustomers," & Format(Now, "dd/mm/yyyy HH:NN")
            SVRMsg.Body = ""
            SVRMsg.Send QSVR
        Next j
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClearCustomersOnFD"
End Sub
Public Sub ClearProductsOnFD()
    On Error GoTo errHandler
Dim j As Integer

        For j = 1 To UBound(arQueuePath)
            QI.FormatName = "DIRECT=OS:" & arQueuePath(j)
            Set QSVR = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            Set SVRMsg = New MSMQMessage
            SVRMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
            SVRMsg.Label = "ClearProducts," & Format(Now, "dd/mm/yyyy HH:NN")
            SVRMsg.Body = ""
            SVRMsg.Send QSVR
        Next j
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClearProductsOnFD"
End Sub
Private Sub Class_Initialize()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim i As Integer
Dim OpenResult As Integer
'
'-------------------------------
    OpenResult = oPC.OpenDBShort
'-------------------------------
'    Set QI = New MSMQQueueInfo
'
    'Prepare queue path array
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tPOSCLIENT", oPC.COShort, adOpenForwardOnly, adLockOptimistic
    i = 1
    Do While Not rs.EOF
        ReDim Preserve arQueuePath(i)
        arQueuePath(i) = rs!MachineName & "\Private$\QSVR"
        rs.MoveNext
        i = i + 1
    Loop
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    'Set up receiving queue (always local queue)
    Set QI = New MSMQQueueInfo
    QI.pathname = oPC.NameOfPC & "\Private$\" & "QPOS"
    On Error Resume Next
    QI.Create
    On Error GoTo errHandler
    Set QPOS = QI.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
    Set POSEvent = New MSMQEvent
    QPOS.EnableNotification POSEvent
        
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.Class_Initialize", , EA_NORERAISE
    HandleError
End Sub



Private Sub PlaceInQueue(pRs As ADODB.Recordset, strLabel As String, strTableName As String)
    On Error GoTo errHandler
Dim j As Integer
Dim OpenResult As Integer

    If Not pRs.EOF Then
        For j = 1 To UBound(arQueuePath)
            QI.FormatName = "DIRECT=OS:" & arQueuePath(j)
            Set QSVR = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            Set SVRMsg = New MSMQMessage
            SVRMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
            SVRMsg.Label = strLabel & "," & Format(Now, "dd/mm/yyyy HH:NN")
            SVRMsg.Body = pRs
            SVRMsg.Send QSVR
        Next j

        oPC.COShort.CommandTimeout = 0
        Select Case strTableName
        Case "tProdUpdates"
            oPC.COShort.Execute "DELETE tProdupdates FROM (SELECT * FROM #PRODS) as t1 WHERE tPRODUPDATES.PRU_GUID = t1.PRU_GUID"
        Case "tStaffMemberUpdate"
            oPC.COShort.Execute "DELETE tStaffMemberUpdate"
        Case "tTPUpdate_Cust"
            oPC.COShort.Execute "DELETE tTPUpdate_Cust FROM (SELECT TOP 2000 * FROM tTPUpdate_Cust Order BY CU_TriggerDate) as t1 WHERE tTPUpdate_Cust.CU_GUID = t1.CU_GUID"
        Case "tCOUpdate"
            oPC.COShort.Execute "DELETE  FROM  tCOUpdate "
        Case "tMarketing_Changes"
            oPC.COShort.Execute "DELETE tMarketing_Changes"
        Case "tAPPUpdate"
            oPC.COShort.Execute "DELETE from tAPPUpdate "
        End Select
            RaiseEvent DataSent
    End If
    If pRs.State = adStateOpen Then pRs.Close
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.PlaceInQueue(pRs,strLabel,strTableName)", Array(pRs, strLabel, _
         strTableName)
End Sub


Public Property Get DBName() As String
    On Error GoTo errHandler
    DBName = "DATABASE NAME"
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.DBName"
End Property

Sub POSEvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
Dim rs As New ADODB.Recordset
Dim lngResult As Integer
Dim QPOS As MSMQQueue
Dim POSmsg As MSMQMessage
Dim QIAck As MSMQQueueInfo
Dim QPOSAck As MSMQQueue
Dim POSAckMsg As MSMQMessage
Dim oStream As ADODB.Stream
Dim strMessage As String
    On Error GoTo errHandler

    Set QPOS = Queue
    Set POSmsg = QPOS.Receive(ReceiveTimeOut:=0)
    If Not (POSmsg Is Nothing) Then
    
        If POSmsg.Label = "STOPPOLLING" Then
            RaiseEvent SwitchPolling(False)
            QPOS.EnableNotification POSEvent
            Exit Sub
        ElseIf POSmsg.Label = "STARTPOLLING" Then
            RaiseEvent SwitchPolling(True)
            QPOS.EnableNotification POSEvent
            Exit Sub
        End If
        strMessage = POSmsg.Body
        SaveSaleData strMessage
        RaiseEvent FileProcessed(strTillpoint)
        
        'Send response if ordinary sale (not XEND or ZEND)
        If Left(POSmsg.Label, 3) = "POS" Then
            Set QIAck = POSmsg.ResponseQueueInfo
            If Not QIAck Is Nothing Then
                Set QPOSAck = QIAck.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
                Set POSAckMsg = New MSMQMessage
                POSAckMsg.Label = "Ack:" & POSmsg.Label
                POSAckMsg.Body = ""
                POSAckMsg.CorrelationId = POSmsg.ID
                POSAckMsg.Send QPOSAck
                QPOSAck.Close
            End If
        End If
    End If
    QPOS.EnableNotification POSEvent
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.POSEvent_Arrived"
End Sub

Private Function SaveSaleData(pSalesMessage As String)
    On Error GoTo errHandler
    
Dim msg As String
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim strUpdate As String
Dim bOK As Integer
Dim bMadeChange As Boolean
Dim XMLDOC As New ujXML
Dim OpenResult As Integer
Dim strExchangePart As String
Dim strCSLPart As String
Dim strPaymentPart As String
Dim tmp() As String

    tmp = Split(pSalesMessage, "|")
    strExchangePart = tmp(0)
    If UBound(tmp) > 0 Then
        strCSLPart = tmp(1)
    End If
    If UBound(tmp) > 1 Then
        strPaymentPart = tmp(2)
    End If
    tmp = Split(strExchangePart, vbTab)
    
    strType = tmp(0)
    strZID = tmp(1)
    strTillpoint = FNS(tmp(2))
    dteZStartDate = FND(tmp(3))
    dteZEndDate = FND(tmp(4))
    If strType <> "Z" Then
        dteZNominalDate = FND(tmp(5))
        strOPID = FNS(tmp(6))
        dteOPStartDate = FND(tmp(7))
        dteOPEndDate = FND(tmp(8))
    End If
    If strType = "E" Then
        lngOperatorID = FNN(tmp(9))
        lngSupervisorID = FNN(tmp(10))
        strEXCHID = FNS(tmp(11))
        dteTimeOfSale = FND(tmp(12))
        lngExchOperatorID = FNN(tmp(13))
        lngExchangeNumber = FNN(tmp(14))
        lngSaleValue = FNN(tmp(15))
        lngDiscountValue = FNN(tmp(16))
        lngVATValue = FNN(tmp(17))
        lngChangeGiven = FNN(tmp(18))
        lngLoyaltyValue = FNN(tmp(19))
        strExchType = FNS(tmp(20))
        strNote = Left(FNS(tmp(21)), 100)
        lngVoids = FNN(tmp(22))
        lngTPID = FNN(tmp(23))
    End If
   
    bOK = 99

'-------------------------------
    OpenResult = oPC.OpenDBShort
'-------------------------------
    oPC.COShort.BeginTrans
    'Update or add to tZSession
    'Update or add to tCS
    'Update or add to tOPSession
    'Add to tExchange
    Select Case strType
    Case "O"
        strSQL = "UPDATE tCS SET CS_ENDSession = '" & ReverseDateTime(dteOPEndDate) & "' WHERE CS_GUID = '" & strOPID & "'"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandText = strSQL
        cmd.CommandType = adCmdText
        cmd.Execute
    Case "Z"
        strSQL = "UPDATE tZSession SET Z_ENDDATE = '" & ReverseDateTime(dteZEndDate) & "' WHERE Z_ID = '" & strZID & "'"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandText = strSQL
        cmd.CommandType = adCmdText
        cmd.Execute
    Case "E"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandText = "sp_UPDATEPOS_1"
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        
        Set prm = cmd.CreateParameter("@ZID", adGUID, adParamInput, , strZID)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@STARTDATE", adDate, adParamInput, , dteZStartDate)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@ENDDATE", adDate, adParamInput, , dteZEndDate)
        cmd.Parameters.Append prm
        Set prm = Nothing
                    
        Set prm = cmd.CreateParameter("@NOMINALDATE", adDate, adParamInput, , dteZNominalDate)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@TILLPOINT", adVarChar, adParamInput, 20, strTillpoint)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@OPSID", adGUID, adParamInput, , strOPID)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@OPOPERATORID", adInteger, adParamInput, , lngOperatorID)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@OPSSTARTTIME", adDate, adParamInput, , dteOPStartDate)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@OPSENDTIME", adDate, adParamInput, , dteOPEndDate)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@EXCHOPERATORID", adInteger, adParamInput, , lngExchOperatorID)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@EXCHSUPERVISORID", adInteger, adParamInput, , lngExchOperatorID)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@EXCHNUMBER", adInteger, adParamInput, , lngExchangeNumber)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , strEXCHID)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@EXZSESSID", adGUID, adParamInput, , strZID)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@EXOPSESSID", adGUID, adParamInput, , strOPID)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@EXCHDATE", adDate, adParamInput, , dteTimeOfSale)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@EXCHSALESVALUE", adInteger, adParamInput, , lngSaleValue)
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = cmd.CreateParameter("@EXCHDISCOUNTVALUE", adInteger, adParamInput, , lngDiscountValue)
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = cmd.CreateParameter("@EXCHVATVALUE", adInteger, adParamInput, , lngVATValue)
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = cmd.CreateParameter("@EXCHCHANGEGIVEN", adInteger, adParamInput, , lngChangeGiven)
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = cmd.CreateParameter("@EXCHLOYALTYVALUE", adInteger, adParamInput, , lngLoyaltyValue)
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = cmd.CreateParameter("@EXCHTYPE", adVarChar, adParamInput, 10, strExchType)
        cmd.Parameters.Append prm
        Set prm = Nothing

        Set prm = cmd.CreateParameter("@EXCHVOIDS", adInteger, adParamInput, , lngVoids)
        cmd.Parameters.Append prm
        Set prm = Nothing
              
        Set prm = cmd.CreateParameter("@EXCHCUSTOMERNAME", adVarChar, adParamInput, 100, strNote)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@EXCHTPID", adInteger, adParamInput, , IIf(lngTPID = 0, oPC.CustomerTPID, lngTPID))
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@TRID", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        Set prm = Nothing
        
        Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        Set prm = cmd.CreateParameter("@ErrPOSITION", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        Set prm = Nothing
       
        cmd.Execute
        lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
        lngErrPos = FNN(cmd.Parameters("@ErrPOSITION"))
        If lngErrstatus <> 0 Then
            LogSaveToFile "Error in Stored procedure sp_UPDATEPOS_1 Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@ERRPOSITION"))
            GoTo ErrHandlerControl
        End If
        
        TRID = IIf(IsNull(cmd.Parameters("@TRID")), 0, cmd.Parameters("@TRID"))
        
        oPC.COShort.Execute "DELETE FROM tTMPBOOKCLUB"
    
    '====================================
        ProcessCSLs strCSLPart
    '====================================
        
    'Consolidate lines if necessary
        If strExchType = "APP" Then ConsolidateApproLines TRID, bMadeChange

SKIPCSL:
    '====================================
        ProcessPayments strPaymentPart
    '====================================
        
        'If this exchange is set to void another then we must do so
        If lngVoids > 0 Then VoidAnExchange strZID, lngVoids
        
        Set cmd = Nothing
        
        If strExchType = "AR" Then AUTOGenerateReturnAndInvoice strEXCHID 'takes lines inserted into tTmpBookclub and uses them
                                                                          'to create invoice and return
                                                                          'associates the invoice with an exchange
    End Select
    
    bOK = True
    
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
    
ErrHandlerControl:
    oPC.COShort.RollbackTrans
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.SaveSaleData(pSalesMessage)", Array(pSalesMessage)
End Function


Private Sub ProcessPayments(pPaymentPart As String)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim strUpdate As String
'PaymentLines
Dim strPaymentType As String
Dim lngAmt As Long
Dim strRefCode As String
Dim strPaymentNote As String
Dim lngPaymentCOLID As Long
Dim i As Integer
Dim tmp() As String
Dim lin() As String
    tmp = Split(pPaymentPart, "~")
    For i = 0 To UBound(tmp)
        lin = Split(tmp(i), vbTab)
        
        strPaymentType = FNS(lin(0))
        lngAmt = FNN(lin(1))
        strRefCode = FNS(lin(2))
        strPaymentNote = FNS(lin(3))
        lngPaymentCOLID = FNN(lin(4))

         Set cmd = Nothing
         Set cmd = New ADODB.Command
         cmd.ActiveConnection = oPC.COShort
         cmd.CommandType = adCmdStoredProc
         cmd.NamedParameters = True
         cmd.CommandText = "sp_UPDATEPOS_2"
         Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , strEXCHID)
         cmd.Parameters.Append prm
         Set prm = Nothing
    
         Set prm = cmd.CreateParameter("@PAYTYPE", adChar, adParamInput, 4, strPaymentType)
         cmd.Parameters.Append prm
         Set prm = Nothing
    
         Set prm = cmd.CreateParameter("@PAYAMT", adInteger, adParamInput, , IIf(FNS(strExchType) = "PCC", lngAmt * -1, lngAmt))
         cmd.Parameters.Append prm
         Set prm = Nothing
    
         Set prm = cmd.CreateParameter("@PAYREF", adVarChar, adParamInput, 100, strRefCode)
         cmd.Parameters.Append prm
         Set prm = Nothing
         Set prm = cmd.CreateParameter("@PAYNOTE", adVarChar, adParamInput, 50, strPaymentNote)
         cmd.Parameters.Append prm
         Set prm = Nothing
    
    
         Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
         cmd.Parameters.Append prm
         Set prm = Nothing
    
         Set prm = cmd.CreateParameter("@ErrPOSITION", adInteger, adParamOutput)
         cmd.Parameters.Append prm
         Set prm = Nothing
    
         cmd.Execute
         lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
    
         If lngErrstatus <> 0 Then
             LogSaveToFile "Error in Stored procedure sp_UPDATEPOS_2, Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@ERRPOSITION"))
             GoTo errHandler
         End If
    
         If strPaymentType = "RD" Then  'Deposit is redeemed - must mark C.O.L deposit status to 'RD' (redeemed)]
             strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'E' WHERE COL_ID = " & lngPaymentCOLID
             oPC.COShort.Execute strUpdate
         ElseIf strPaymentType = "RDEP" Then 'Deposit is refunded
             strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'X' WHERE COL_ID = " & lngPaymentCOLID
             oPC.COShort.Execute strUpdate
         End If
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ProcessPayments(pPaymentPart)", pPaymentPart
End Sub
Public Sub ProcessCSLs(pCSLPart As String)
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim strUpdate As String

'CSL fields
Dim strPID As String
Dim lngCOLID As Long
Dim lngQty As Long
Dim lngPrice As Long
Dim lngPriceAlteration As Long
Dim lngDiscount As Long
Dim dblDiscountRate As Double
Dim dblVATRATE As Double
Dim strCounterfoil As String
Dim strDiscountDescription As String
Dim tmp() As String
Dim lin() As String
Dim i As Integer

    tmp = Split(pCSLPart, "~")
    For i = 0 To UBound(tmp)
        lin = Split(tmp(i), vbTab)
        
        strPID = FNS(lin(0))
        lngCOLID = FNN(lin(1))
        lngQty = FNN(lin(2))
        lngPrice = FNN(lin(3))
        lngPriceAlteration = FNN(lin(4))
        'MsgBox "problem with price alteraion"
        lngDiscount = FNN(lin(5))
        dblDiscountRate = FNDBL(lin(6))
        dblVATRATE = FNDBL(lin(7))
        strCounterfoil = FNS(lin(8))
        strDiscountDescription = FNS(lin(9))
    ''
        Set cmd = Nothing
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandType = adCmdStoredProc
        cmd.NamedParameters = True
        If strExchType = "APP" Then
            cmd.CommandText = "sp_APP"
    
        ElseIf strExchType = "AR" Then
            cmd.CommandText = "sp_APPRET"
            Set prm = cmd.CreateParameter("@APPLID", adInteger, adParamInput, , lngCOLID)
            cmd.Parameters.Append prm
            Set prm = Nothing
        ElseIf strExchType <> "DEP" And strExchType <> "RDEP" And strExchType <> "OR" Then
            cmd.CommandText = "sp_CSL"
        Else
            If strExchType = "DEP" Then  'Deposit is accepted for order
                strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'P' WHERE COL_ID = " & lngCOLID
                oPC.COShort.Execute strUpdate
            End If
            If strExchType = "RDEP" Then  'Deposit is accepted for order
                strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'X' WHERE COL_ID = " & lngCOLID
                oPC.COShort.Execute strUpdate
            End If
        End If
    ''
        Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@PID", adGUID, adParamInput, , strPID)
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@Qty", adInteger, adParamInput, , lngQty)
        cmd.Parameters.Append prm
        Set prm = Nothing
    ''
        Set prm = cmd.CreateParameter("@Price", adInteger, adParamInput, , lngPrice)
        cmd.Parameters.Append prm
        Set prm = Nothing
        Set prm = cmd.CreateParameter("@PriceAlteration", adInteger, adParamInput, , 0)
        cmd.Parameters.Append prm
        Set prm = Nothing
       ' MsgBox "Problem with price alteration"
        Set prm = cmd.CreateParameter("@DiscountRate", adNumeric, adParamInput, , dblDiscountRate)
        prm.NumericScale = 2
        prm.Precision = 8
        cmd.Parameters.Append prm
        Set prm = Nothing
    ''
        Set prm = cmd.CreateParameter("@Discount", adInteger, adParamInput, , lngDiscount)
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@VRate", adInteger, adParamInput, , dblVATRATE)
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@STID", adInteger, adParamInput, , oPC.DefaultStoreID)
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , IIf(lngTPID = 0, oPC.CustomerTPID, lngTPID))
        cmd.Parameters.Append prm
        Set prm = Nothing
    ''
        Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , strEXCHID)
        cmd.Parameters.Append prm
        Set prm = Nothing
    '
        Set prm = cmd.CreateParameter("@COUNTERFOIL", adVarChar, adParamInput, 30, strCounterfoil)
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@DISCOUNTDESCRIPTION", adVarChar, adParamInput, 100, strDiscountDescription)
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@ID", adInteger, adParamOutput) 'dummy in this context
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        Set prm = Nothing
    
        Set prm = cmd.CreateParameter("@pPOSITION", adInteger, adParamOutput)
        cmd.Parameters.Append prm
        Set prm = Nothing
    ''
        cmd.Execute
        lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
    ''
        If lngErrstatus <> 0 Then
            LogSaveToFile "Error in Stored procedure " & cmd.CommandText & " Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@pPOSITION"))
            GoTo ErrHandlerControl
        End If
    Next i
    Exit Sub
    
ErrHandlerControl:
    oPC.COShort.RollbackTrans
    Exit Sub

End Sub
Private Sub VoidAnExchange(pZID As String, pExchNum As Long)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandType = adCmdStoredProc
    cmd.NamedParameters = True
    cmd.CommandText = "sp_VOIDEXCHANGE"
    
    Set prm = cmd.CreateParameter("@ZID", adGUID, adParamInput, , pZID)
    cmd.Parameters.Append prm
    Set prm = Nothing

    Set prm = cmd.CreateParameter("@EXCHNUMBER", adInteger, adParamInput, , pExchNum)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    cmd.Execute
    Set cmd = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.VoidAnExchange(pZID,pExchNum)", Array(pZID, pExchNum)
End Sub

Public Sub AUTOGenerateReturnAndInvoice(pEXCHID As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_APPRO_INV_AND_RETURN_POS"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , pEXCHID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@TRSTATUS", adInteger, adParamInput, , 3)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@IID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@ARID", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing

    Set par = cmd.CreateParameter("@POSITION", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@ERR", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.Execute
    
    Set cmd = Nothing



    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.AUTOGenerateReturnAndInvoice(pEXCHID)", pEXCHID
End Sub

Public Sub ConsolidateApproLines(TRID As Long, bMadeChange As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long

    Set cmd = New ADODB.Command
    cmd.CommandText = "ConsolidateApproLines"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("Return", adInteger, adParamReturnValue)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    
    cmd.Execute
    iReturn = cmd.Parameters("Return")
    If iReturn > 1 Then Err.Raise EXC_SP_FAILED, "a_APP_P:ConsolidateLines", "SP failed"
    bMadeChange = (iReturn = 1)
    
    Set cmd = Nothing

Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_APP_P.ConsolidateApproLines_Svr(TRID,bMadeChange)", Array(TRID, bMadeChange)
End Sub

