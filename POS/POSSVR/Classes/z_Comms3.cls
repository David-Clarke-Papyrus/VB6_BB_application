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
'ption Base 1

'=================================================
'POSExchange variables from XML document
'ZSession
Dim strType As String
Dim strZID As String
Dim strTillpoint As String
Dim dteZStartDate As String
Dim dteZEndDate As String
Dim dteZNominalDate As String
'OPSession
Dim dteOPStartDate As String
Dim dteOPEndDate As String
Dim strOPID As String
Dim lngOperatorID As Long
Dim lngSupervisorID As Long
'Exchange
Dim strEXCHID As String
Dim lngExchOperatorID As Long
Dim dteTimeOfSale As String
Dim lngSaleValue As Long
Dim lngDiscountValue As Long
Dim lngVATValue As Long
Dim lngChangeGiven As Long
Dim lngLoyaltyValue As Long
Dim strExchType As String
Dim strNote As String
Dim lngVoids As Long
Dim lngTPID As Long
Dim lngSalesRepID As Long
Dim dblFloat As Double
Dim sFloatBreakdown As String
Dim lngExchangeNumber As Long

Dim TRID As Long
Dim lngErrstatus As Long
Dim lngErrPos As Long
Dim lngLastPayment As Long
Dim fs As New FileSystemObject
'=================================================
Dim strPos As String
Dim arQueuePath() As String

Public Event FileProcessed(Till As String) '0 = Other, 1 = Sale, 2 = Update
Public Event DataSent()
Public Event DBStatusMsg(val As String)
'Turns off POS_SErver while gateway fetches data, triggers remain active and data is sent l
'Public Event SwitchPolling(bOn As Boolean)
Dim bDBOK As Boolean

Private WithEvents POSEvent As MSMQEvent
Attribute POSEvent.VB_VarHelpID = -1
Private QPOS As MSMQQueue
Private WithEvents SVREvent As MSMQEvent
Attribute SVREvent.VB_VarHelpID = -1
Private QSVR As MSMQQueue

Dim qi As MSMQQueueInfo
Dim QSVRAck As MSMQQueue
Dim SVRMsg As msmqMessage
'Dim msmqMessage As System.Messaging.Message
Event ErrorRegistering(pStationName As String)
Event ClientStatusChange()
Event QPOSTriggerStatus(pOn As Boolean)
Event POSMessageReceived(pStationName As String)
Event DBStatus(val As Boolean)
Dim lngErrorCount As Long


Public Function PutMasterDataInOutbox() As Boolean
10        On Error GoTo errHandler
Dim OpenResult As Integer
Dim cnt As Long
Dim strPos As String
Dim j
Dim T1 As Boolean
Dim t2 As Boolean
Dim t3 As Boolean
Dim t4 As Boolean
Dim t5 As Boolean
Dim t6 As Boolean
Dim t7 As Boolean
Dim reslt As Long

Dim oTF As New z_TextFile
Dim SVRRS As ADODB.Recordset

20        On Error Resume Next
30        j = UBound(arQueuePath)
40        If Err Then
50            PutMasterDataInOutbox = False
60            Exit Function
70        End If
80        On Error GoTo errHandler
        '  oTF.OpenTextFileToAppend oPC.LocalRoot & "\UpdatesFromServerToPOS.txt"
        '  SetQPOSTrigger
90   ' Err.Raise 12345
      '-------------------------------
100       OpenResult = oPC.OpenDBSHort
      '-------------------------------
110      RaiseEvent DBStatusMsg("Checking for server updates")
120     MsgWaitObj 300
130       If OpenResult = 2 Then
140           RaiseEvent DBStatus(False)
150           PutMasterDataInOutbox = False
160           GoTo EXIT_Handler
170       Else
180           RaiseEvent DBStatus(True)
190       End If
200       CheckProductUpdatesAvailable T1, t2, t3, t4, t5, t6, t7, reslt
210       If reslt = 0 Then
220         If T1 + t2 + t3 + t4 + t5 + t6 + t7 > 0 Then
230             RaiseEvent DBStatusMsg("Updating POS stations")
240         Else
250             RaiseEvent DBStatusMsg("Idle")
260         End If
270       ElseIf reslt = 1 Then
280         RaiseEvent DBStatusMsg("Database changes are not being detected. If this message repeats frequently please contact support.")
290         MsgWaitObj 300
300       ElseIf reslt = 2 Then
310         RaiseEvent DBStatusMsg("Database changes are not being detected. If this message repeats frequently please contact support.")
320         MsgWaitObj 300
330       End If
          
340       oPC.COShort.CommandTimeout = 1800                'allow 30 minutes for command to execute
350       If T1 Then
360           oPC.COShort.Execute "SELECT TOP 2000 tAPPRLUPDATE.* into #APPRL FROM tAPPRLUPDATE"
370           Set SVRRS = GetRecordset("SELECT * FROM #APPRL")
380           If bDebug Then LogSaveToFile Format(Now(), "dd/mm/yy HH:NN") & "SELECT * FROM #APPRL: " & CStr(SVRRS.RecordCount)
390           PlaceInQueue SVRRS, "APPRL", "tAPPRLUPDATE"
400           oPC.COShort.Execute "DROP table #APPRL"
410           Set SVRRS = Nothing
420       End If
430       If t2 Then
440           oPC.COShort.Execute "SELECT TOP 2000 tAPPUpdate.* into #APPL FROM tAPPUpdate"
450           Set SVRRS = GetRecordset("SELECT * FROM #APPL")
460           If bDebug Then LogSaveToFile Format(Now(), "dd/mm/yy HH:NN") & "SELECT * FROM #APPL: " & CStr(SVRRS.RecordCount)
470           PlaceInQueue SVRRS, "APPL", "tAPPUpdate"
480           oPC.COShort.Execute "DROP table #APPL"
490            Set SVRRS = Nothing
500      End If
510       If t3 Then
520           oPC.COShort.Execute "SELECT TOP 2000 tCOUPDATE.* into #ORDR FROM tCOUPDATE"
530           Set SVRRS = GetRecordset("SELECT * FROM #ORDR")
540           If bDebug Then LogSaveToFile Format(Now(), "dd/mm/yy HH:NN") & "SELECT * FROM #ORDR: " & CStr(SVRRS.RecordCount)
550           PlaceInQueue SVRRS, "ORDR", "tCOUpdate"
560           oPC.COShort.Execute "DROP table #ORDR"
570           Set SVRRS = Nothing
580       End If
590       If t4 Then
600           oPC.COShort.Execute "SELECT TOP 500 tTPUpdate_Cust.* into #CUSTS FROM tTPUpdate_Cust"
610           Set SVRRS = GetRecordset("SELECT * FROM #CUSTS")
620           If bDebug Then LogSaveToFile Format(Now(), "dd/mm/yy HH:NN") & "SELECT * FROM #CUSTS: " & CStr(SVRRS.RecordCount)
630           PlaceInQueue SVRRS, "CUST", "tTPUpdate_Cust"
640           oPC.COShort.Execute "DROP table #CUSTS"
650           Set SVRRS = Nothing
660       End If
670       If t5 Then
680           PlaceInQueue GetRecordset("SELECT * FROM tStaffMemberUpdate"), "STAF", "tStaffMemberUpdate"
690       End If
700       If t6 Then
710           PlaceInQueue GetRecordset("SELECT * FROM tMarketing_Changes"), "MARK", "tMarketing_Changes"
720       End If
730       If t7 Then
740           oPC.COShort.Execute "SELECT TOP 2000 tProdUpdates.* into #PRODS FROM tProdUpdates"
750           Set SVRRS = GetRecordset("SELECT * FROM #PRODS")
760           If bDebug Then LogSaveToFile Format(Now(), "dd/mm/yy HH:NN") & "SELECT * FROM #PRODS: " & CStr(SVRRS.RecordCount)
770           PlaceInQueue SVRRS, "PROD", "tProdUpdates"
780           oPC.COShort.Execute "DROP table #PRODS"
790           Set SVRRS = Nothing
800       End If

810       PutMasterDataInOutbox = True
820       DoEvents
              
EXIT_Handler:
830       On Error Resume Next
      '---------------------------------------------------
840       If OpenResult = 0 Then
850           If Not (oPC Is Nothing) Then oPC.DisconnectDBShort
860       End If
      '---------------------------------------------------
870       Exit Function
errHandler:
880       ErrPreserve
890       lngErrorCount = lngErrorCount + 1
900       RaiseEvent DBStatusMsg("Failing to send master database changes to POS stations. If this message lasts more than 5 minutes report this to your support person.")
910       If lngErrorCount > 20 And lngErrorCount < 999 Then
              lngErrorCount = 999
920           ErrorIn "Z_PollingServices.PutMasterDataInOutbox", , EA_ROLLBACK + EA_RERAISE, oPC.COShort, "strPos", Array(strPos)
930       End If

End Function
Private Function GetRecordset(strSQL As String) As ADODB.Recordset
    On Error GoTo errHandler
Dim rs1 As ADODB.Recordset

        oPC.BeginTransaction
        Set rs1 = New ADODB.Recordset
        rs1.CursorLocation = adUseClient
        rs1.Open strSQL, oPC.COShort, adOpenStatic, adLockOptimistic
        rs1.ActiveConnection = Nothing         'DISCONNECT RECORDSET
        Set GetRecordset = rs1
        Set rs1 = Nothing
        oPC.CommitTransaction

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.GetRecordset(strSQL)", strSQL
End Function
Public Sub ClearOnFD(pInstruction As String)
    On Error GoTo errHandler
Dim j As Integer

        For j = 1 To UBound(arQueuePath)
            qi.FormatName = "DIRECT=OS:" & arQueuePath(j)
            Set QSVR = qi.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            Set SVRMsg = New msmqMessage
            SVRMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
            SVRMsg.Label = pInstruction & "," & Format(Now, "dd/mm/yyyy HH:NN")
            SVRMsg.MaxTimeToReachQueue = 777600   '9 days
            SVRMsg.Body = ""
            SVRMsg.send QSVR
        Next j
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClearCustomersOnFD"
End Sub

Private Sub Class_Initialize()
10        On Error GoTo errHandler
      Dim rs As ADODB.Recordset
      Dim i As Integer
      Dim OpenResult As Integer
      '
        lngErrorCount = 0
      '-------------------------------
20        OpenResult = oPC.OpenDBSHort
      '-------------------------------
30        oPC.Initialize
          'Prepare queue path array
40        Set rs = New ADODB.Recordset
50        rs.Open "SELECT * FROM tPOSCLIENT", oPC.COShort, adOpenForwardOnly, adLockOptimistic
60        i = 1
70        Do While Not rs.EOF
80            ReDim Preserve arQueuePath(i)
90            arQueuePath(i) = rs!MachineName & "\Private$\qsvr"
100           rs.MoveNext
110           i = i + 1
120       Loop
      '---------------------------------------------------
130       If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------
          'create receiving queue (always local queue)
140       Set qi = New MSMQQueueInfo
150       qi.PathName = oPC.NameOfPC & "\Private$\" & "qpos"
160       On Error Resume Next  'probably is already created
170       qi.Create , True
          
180       OpenQPOS
190       Exit Sub
errHandler:
200       If ErrMustStop Then Debug.Assert False: Resume
210       ErrorIn "Z_PollingServices.Class_Initialize", , EA_NORERAISE
220       HandleError2
End Sub

Public Sub OpenQPOS()
    On Error GoTo errHandler
    If QPOS Is Nothing Then
    
        Set qi = New MSMQQueueInfo
        qi.PathName = oPC.NameOfPC & "\Private$\" & "qpos"
        
        On Error Resume Next
        qi.Create , True
        On Error GoTo errHandler
    
        Set QPOS = qi.Open(MQ_RECEIVE_ACCESS, MQ_DENY_NONE)
        Set POSEvent = New MSMQEvent
        SetQPOSTrigger
    End If
    If Not QPOS.IsOpen2 Then
        SetQPOSTrigger
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.OpenQPOS"
End Sub
Public Sub CloseQPOS()
    On Error GoTo errHandler
    If Not QPOS Is Nothing Then
        If QPOS.IsOpen2 Then
            QPOS.Close
            Set QPOS = Nothing
            RaiseEvent QPOSTriggerStatus(False)
        End If
    Else
            RaiseEvent QPOSTriggerStatus(False)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.CloseQPOS"
End Sub

Public Sub SetQPOSTrigger()
On Error Resume Next
   ' MsgBox "Enabling trigger"
    QPOS.EnableNotification POSEvent
    RaiseEvent QPOSTriggerStatus(True)
End Sub


Private Sub PlaceInQueue(pRs As ADODB.Recordset, strLabel As String, strTableName As String)
10        On Error GoTo errHandler
      Dim j As Integer

20        If Not pRs.EOF Then
30            For j = 1 To UBound(arQueuePath)
40                qi.FormatName = "DIRECT=OS:" & arQueuePath(j)
50                Set QSVR = qi.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
60                Set SVRMsg = New msmqMessage
70                SVRMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
80                SVRMsg.Label = strLabel & "," & Format(Now, "dd/mm/yyyy HH:NN")
90                SVRMsg.Body = pRs
100               SVRMsg.MaxTimeToReachQueue = 777600   '9 days
110               SVRMsg.send QSVR
120           Next j

130           oPC.COShort.CommandTimeout = 0
140           Select Case UCase(strTableName)
              Case "TPRODUPDATES"
150               oPC.COShort.Execute "DELETE tProdupdates FROM (SELECT * FROM #PRODS) as t1 WHERE tPRODUPDATES.PRU_GUID = t1.PRU_GUID"
160           Case "TSTAFFMEMBERUPDATE"
170               oPC.COShort.Execute "DELETE tStaffMemberUpdate"
180           Case "TTPUPDATE_CUST"
190               oPC.COShort.Execute "DELETE tTPUpdate_Cust FROM (SELECT * FROM #CUSTS) as t1 WHERE tTPUpdate_Cust.CU_GUID = t1.CU_GUID"
200           Case "TCOUPDATE"
210               oPC.COShort.Execute "DELETE tCOUpdate FROM (SELECT * FROM #ORDR) as t1 WHERE tCOUpdate.COU_GUID = t1.COU_GUID"
220           Case "TMARKETING_CHANGES"
230               oPC.COShort.Execute "DELETE tMarketing_Changes"
240           Case "TAPPUPDATE"
250               oPC.COShort.Execute "DELETE tAPPUpdate FROM (SELECT * FROM #APPL) as t1 WHERE tAPPUpdate.APPL_GUID = t1.APPL_GUID"
260           Case "TAPPRLUPDATE"
270               oPC.COShort.Execute "DELETE tAPPRLUpdate FROM (SELECT * FROM #APPRL) as t1 WHERE tAPPRLUpdate.ARLU_GUID = t1.ARLU_GUID"
280           End Select
290           RaiseEvent DataSent
300       End If
310       If pRs.State = adStateOpen Then pRs.Close
320       Exit Sub
errHandler:
330       If ErrMustStop Then Debug.Assert False: Resume
340       ErrorIn "Z_PollingServices.PlaceInQueue(pRs,strLabel,strTableName)", Array(pRs, strLabel, _
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

Private Sub Class_Terminate()
    Set qi = Nothing
End Sub
Public Sub ReprocessMessage(pMsg As String, strFilename As String)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim iresult As Integer
Dim lQPOS As MSMQQueue
Dim POSmsg As msmqMessage
Dim lQIAck As MSMQQueueInfo
Dim lQPOSAck As MSMQQueue
Dim POSAckMsg As msmqMessage
Dim oStream As ADODB.Stream
Dim strMessage As String
Dim OpenResult As Integer
Dim strEXCHID As String
Dim lngExchangeNumber As Long
Dim MsgName As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    RaiseEvent QPOSTriggerStatus(False)

        strMessage = pMsg
        iresult = SaveSaleData(strMessage)
      '  MsgBox strMessage
        If iresult = 0 And strFilename <> "" Then
            fs.DeleteFile strFilename
        End If
    SetQPOSTrigger
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ReprocessMessage(pMsg,strFileName)", Array(pMsg, strFilename)
End Sub
Sub POSEvent_Arrived(ByVal Queue As Object, ByVal Cursor As Long)
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim iresult As Integer
Dim lQPOS As MSMQQueue
Dim POSmsg As msmqMessage
Dim lQIAck As MSMQQueueInfo
Dim lQPOSAck As MSMQQueue
Dim POSAckMsg As msmqMessage
Dim oStream As ADODB.Stream
Dim strMessage As String
Dim OpenResult As Integer
Dim strEXCHID As String
Dim lngExchangeNumber As Long
Dim MsgName As String

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set lQPOS = Queue
    Set POSmsg = lQPOS.Receive
    RaiseEvent QPOSTriggerStatus(False)
    If Not (POSmsg Is Nothing) Then
        strMessage = POSmsg.Body
        iresult = SaveSaleData(strMessage, strEXCHID, lngExchangeNumber)
'        If iresult = 0 Then
'            fs.DeleteFile MsgName
'        End If
        'Send response if ordinary sale (not XEND or ZEND)
        If Left(POSmsg.Label, 3) = "POS" Then
            Set lQIAck = POSmsg.ResponseQueueInfo
            If Not lQIAck Is Nothing Then
                Set lQPOSAck = lQIAck.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
                Set POSAckMsg = New msmqMessage
                POSAckMsg.Label = "Ack:" & POSmsg.Label
                POSAckMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
                POSAckMsg.Body = strEXCHID
                POSAckMsg.CorrelationId = POSmsg.ID
                POSAckMsg.send lQPOSAck
                lQPOSAck.Close
            End If
        End If
    Else
    '    MsgBox "Pos 2"
    End If
  '  MsgBox "Before Resetting trigger"
    SetQPOSTrigger
    RaiseEvent QPOSTriggerStatus(True)
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.POSEvent_Arrived(Queue,Cursor)", Array(Queue, Cursor), , , "iResult,lngExchangeNumber", Array(iresult, lngExchangeNumber)
    HandleError
End Sub
Private Function LogMessage(msg As String, MsgName As String)
    On Error GoTo errHandler
Dim oTF As New z_TextFile
'    LogSaveToFile msg
'    MsgName = oPC.SharedFolderRoot & "\Exchanges\ex" & Format(Now(), "yymmddHHNNSS") & ".txt"
'    oTF.OpenTextFile MsgName
'    oTF.WriteToTextFile msg
'    oTF.CloseTextFile
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.LogMessage(msg,MsgName)", Array(msg, MsgName)
End Function
Private Function SaveSaleData(pSalesMessage As String, Optional pExchID As String, Optional pExchangeNumber) As Integer
10        On Error GoTo errHandler
          
      Dim msg As String
      Dim cmd As New ADODB.Command
      Dim prm As ADODB.Parameter
      Dim strUpdate As String
      Dim bMadeChange As Boolean
      Dim strExchangePart As String
      Dim strCSLPart As String
      Dim strPaymentPart As String
      Dim tmp() As String
      Dim Result As Integer
      Dim lastExchangeNumber As Long
      Dim appTRID As Long
      Dim strErrPos As String
           If bDebug Then LogSaveToFile " Re-Saving " & pSalesMessage

      'MsgBox "in SaveSaleData"
20        Result = 0
30        tmp = Split(pSalesMessage, "|")
40        strExchangePart = tmp(0)
50        If UBound(tmp) > 0 Then
60            strCSLPart = tmp(1)
70        End If
80        If UBound(tmp) > 1 Then
90            strPaymentPart = tmp(2)
100       End If
110       tmp = Split(strExchangePart, vbTab)
120       strType = tmp(0)
130       strZID = tmp(1)
140       strTillpoint = FNS(tmp(2))
         ' MsgBox "1 " & strTillpoint
150       dteZStartDate = FNS(tmp(3))
160       dteZEndDate = FNS(tmp(4))
170       If strType <> "Z" Then    'And strType <> "O"
180           dteZNominalDate = FNS(tmp(5))
190           strOPID = FNS(tmp(6))
200           dteOPStartDate = FNS(tmp(7))
210           dteOPEndDate = FNS(tmp(8))
220           If UBound(tmp) > 8 Then dblFloat = FNDBL(tmp(9))
230           If UBound(tmp) > 9 Then sFloatBreakdown = FNS(tmp(10))
240       Else
250           lastExchangeNumber = FNN(tmp(5))
260       End If
270       If strType = "E" Or strType = "A" Or strType = "P" Or strType = "CN" Then
280           RaiseEvent POSMessageReceived(strTillpoint)
290           lngOperatorID = FNN(tmp(9))
300           lngSupervisorID = FNN(tmp(10))
310           strEXCHID = FNS(tmp(11))
320           pExchID = strEXCHID
330           dteTimeOfSale = FNS(tmp(12))
340           lngExchOperatorID = FNN(tmp(13))
350           lngExchangeNumber = FNN(tmp(14))
360           pExchangeNumber = lngExchangeNumber
370           lngSaleValue = FNN(tmp(15))
380           lngDiscountValue = FNN(tmp(16))
390           lngVATValue = FNN(tmp(17))
400           lngChangeGiven = FNN(tmp(18))
410           lngLoyaltyValue = FNN(tmp(19))
420           strExchType = FNS(tmp(20))
430           strNote = Left(FNS(tmp(21)), 5000)
440           lngVoids = FNN(tmp(22))
450           lngTPID = FNN(tmp(23))
              ' THis next was necessary because of the transition period where some stores had this next field others not
460           If UBound(tmp) > 23 Then
470               lngSalesRepID = FNN(tmp(24))
480           Else
490               lngSalesRepID = 0
500           End If
510       End If
520       RaiseEvent DBStatusMsg("Posting exchange")
530      If bDebug Then LogSaveToFile "Pos 1 " & strType
          'Update or add to tZSession
          'Update or add to tCS
          'Update or add to tOPSession
          'Add to tExchange
540       Select Case strType
          Case "O"
550           strSQL = "UPDATE tCS SET CS_ENDSession = '" & dteOPEndDate & "' , CS_FLOATVALUE = " & CStr(dblFloat) & ",CS_FLOATBREAKDOWNAtSTART = '" & sFloatBreakdown & "' WHERE CS_GUID = '" & strOPID & "'"
560           Set cmd = New ADODB.Command
570           cmd.ActiveConnection = oPC.COShort
580           cmd.CommandText = strSQL
590           cmd.CommandType = adCmdText
600           cmd.Execute
        GoTo SkipforZENDXEND:
610       Case "Z"
620           strSQL = "UPDATE tZSession SET Z_ENDDATE = '" & dteZEndDate & "',Z_LastExchangeNumber = " & lastExchangeNumber & " WHERE Z_ID = '" & strZID & "'"
630           Set cmd = New ADODB.Command
640           cmd.ActiveConnection = oPC.COShort
650           cmd.CommandText = strSQL
660           cmd.CommandType = adCmdText
670           cmd.Execute
        GoTo SkipforZENDXEND:
680       Case "E", "A", "P", "CN"
690           strErrPos = "NOT in tran"
            ''  oPC.BeginTransaction
700           strErrPos = "TA count " & CStr(oPC.TransactionCount)
              
710           Set cmd = New ADODB.Command
720           cmd.ActiveConnection = oPC.COShort
              'This sp inserts the Exchange record and ensures it has a tCS and a tZSession to belong to
730           cmd.CommandText = "sp_UPDATEPOS_1"
740           cmd.CommandType = adCmdStoredProc
750           cmd.NamedParameters = True
              
760           Set prm = cmd.CreateParameter("@ZID", adGUID, adParamInput, , strZID)
770           cmd.Parameters.Append prm
780           Set prm = Nothing
              
790           Set prm = cmd.CreateParameter("@STARTDATE", adVarChar, adParamInput, 25, dteZStartDate)
800           cmd.Parameters.Append prm
810           Set prm = Nothing
              
820           Set prm = cmd.CreateParameter("@ENDDATE", adVarChar, adParamInput, 25, dteZEndDate)
830           cmd.Parameters.Append prm
840           Set prm = Nothing
                          
850           Set prm = cmd.CreateParameter("@NOMINALDATE", adVarChar, adParamInput, 25, dteZNominalDate)
860           cmd.Parameters.Append prm
870           Set prm = Nothing
              
880           Set prm = cmd.CreateParameter("@TILLPOINT", adVarChar, adParamInput, 15, strTillpoint)
890           cmd.Parameters.Append prm
900           Set prm = Nothing
             
910           Set prm = cmd.CreateParameter("@OPSID", adGUID, adParamInput, , strOPID)
920           cmd.Parameters.Append prm
930           Set prm = Nothing
             
940           Set prm = cmd.CreateParameter("@OPOPERATORID", adInteger, adParamInput, , lngOperatorID)
950           cmd.Parameters.Append prm
960           Set prm = Nothing
             
970           Set prm = cmd.CreateParameter("@OPSSTARTTIME", adVarChar, adParamInput, 25, dteOPStartDate)
980           cmd.Parameters.Append prm
990           Set prm = Nothing
             
1000          Set prm = cmd.CreateParameter("@OPSENDTIME", adVarChar, adParamInput, 25, dteOPEndDate)
1010          cmd.Parameters.Append prm
1020          Set prm = Nothing
              
1030          Set prm = cmd.CreateParameter("@EXCHOPERATORID", adInteger, adParamInput, , lngExchOperatorID)
1040          cmd.Parameters.Append prm
1050          Set prm = Nothing
              
1060          Set prm = cmd.CreateParameter("@EXCHSUPERVISORID", adInteger, adParamInput, , lngSupervisorID)
1070          cmd.Parameters.Append prm
1080          Set prm = Nothing
              
1090          Set prm = cmd.CreateParameter("@EXCHNUMBER", adInteger, adParamInput, , lngExchangeNumber)
1100          cmd.Parameters.Append prm
1110          Set prm = Nothing
              
1120          Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , strEXCHID)
1130          cmd.Parameters.Append prm
1140          Set prm = Nothing
             
1150          Set prm = cmd.CreateParameter("@EXZSESSID", adGUID, adParamInput, , strZID)
1160          cmd.Parameters.Append prm
1170          Set prm = Nothing
             
1180          Set prm = cmd.CreateParameter("@EXOPSESSID", adGUID, adParamInput, , strOPID)
1190          cmd.Parameters.Append prm
1200          Set prm = Nothing
              'MsgBox "dteTimeOfSale =" & dteTimeOfSale
1210          Set prm = cmd.CreateParameter("@EXCHDATE", adVarChar, adParamInput, 25, dteTimeOfSale)
1220          cmd.Parameters.Append prm
1230          Set prm = Nothing
             
1240          Set prm = cmd.CreateParameter("@EXCHSALESVALUE", adInteger, adParamInput, , lngSaleValue)
1250          cmd.Parameters.Append prm
1260          Set prm = Nothing
1270          Set prm = cmd.CreateParameter("@EXCHDISCOUNTVALUE", adInteger, adParamInput, , lngDiscountValue)
1280          cmd.Parameters.Append prm
1290          Set prm = Nothing
1300          Set prm = cmd.CreateParameter("@EXCHVATVALUE", adInteger, adParamInput, , lngVATValue)
1310          cmd.Parameters.Append prm
1320          Set prm = Nothing
1330          Set prm = cmd.CreateParameter("@EXCHCHANGEGIVEN", adInteger, adParamInput, , lngChangeGiven)
1340          cmd.Parameters.Append prm
1350          Set prm = Nothing
1360          Set prm = cmd.CreateParameter("@EXCHLOYALTYVALUE", adInteger, adParamInput, , lngLoyaltyValue)
1370          cmd.Parameters.Append prm
1380          Set prm = Nothing
1390          Set prm = cmd.CreateParameter("@EXCHTYPE", adVarChar, adParamInput, 10, strExchType)
1400          cmd.Parameters.Append prm
1410          Set prm = Nothing

1420          Set prm = cmd.CreateParameter("@EXCHVOIDS", adInteger, adParamInput, , lngVoids)
1430          cmd.Parameters.Append prm
1440          Set prm = Nothing
                   ' MsgBox "strNote = " & strNote
1450          Set prm = cmd.CreateParameter("@EXCHCUSTOMERNAME", adVarChar, adParamInput, 5000, strNote)
1460          cmd.Parameters.Append prm
1470          Set prm = Nothing
             
1480          Set prm = cmd.CreateParameter("@EXCHTPID", adInteger, adParamInput, , IIf(lngTPID = 0, oPC.CustomerTPID, lngTPID))
1490          cmd.Parameters.Append prm
1500          Set prm = Nothing
              
1510          Set prm = cmd.CreateParameter("@EXCHSalesRepID", adInteger, adParamInput, , lngSalesRepID)
1520          cmd.Parameters.Append prm
1530          Set prm = Nothing
              
1540          Set prm = cmd.CreateParameter("@TRID", adInteger, adParamOutput)
1550          cmd.Parameters.Append prm
1560          Set prm = Nothing
              
1570          Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
1580          cmd.Parameters.Append prm
1590          Set prm = Nothing
             
1600          Set prm = cmd.CreateParameter("@ErrPOSITION", adInteger, adParamOutput)
1610          cmd.Parameters.Append prm
1620          Set prm = Nothing
1630          cmd.CommandTimeout = 180
1640                  If bDebug Then LogSaveToFile "Pos Before UPDATEPOS1: ExchID = " & pExchID
1650          cmd.Execute
1660          lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
1670          lngErrPos = FNN(cmd.Parameters("@ErrPOSITION"))
1680          If lngErrstatus <> 0 Then
1690              LogSaveToFile "Error in Stored procedure sp_UPDATEPOS_1 Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@ERRPOSITION"))
1700              GoTo ErrHandlerControl
1710          End If
1720          TRID = IIf(IsNull(cmd.Parameters("@TRID")), 0, cmd.Parameters("@TRID"))
1730          oPC.COShort.CommandTimeout = 180
              
              
1740          oPC.COShort.Execute "Truncate Table tTemporaryDetailLines"
1750          Set cmd = Nothing
              
              
          '====================================
              'this next ensures no duplicates are retained
1760              If bDebug Then LogSaveToFile "Pos Before CSLs: ExchID = " & pExchID
1770          oPC.COShort.Execute "DELETE FROM tCSL WHERE CSL_EXCHANGE_GUID = '" & pExchID & "'"
              
          'handle each detail line
1780          Result = ProcessCSLs(strCSLPart, dteTimeOfSale)
1790          If Result <> 0 Then
1800              GoTo ErrHandlerControl
1810          End If
                  
SKIPCSL:
          '====================================
              'this next ensures no duplicates are retained
1820              If bDebug Then LogSaveToFile "Pos Before tPAYMENT: Number = " & CStr(pExchangeNumber)
1830          oPC.COShort.Execute "DELETE FROM tPAYMENT WHERE PAY_EXCHANGE_GUID =  '" & pExchID & "'"
          
          'handle each payment lines
1840          Result = ProcessPayments(strPaymentPart)
1850          If Result <> 0 Then
1860              GoTo ErrHandlerControl
1870          End If
          '====================================
              'If this exchange is set to void another then we must do so
1880          If lngVoids > 0 Then VoidAnExchange strZID, lngVoids
1890          Set cmd = Nothing
              'StrType: E is cash exchange,
              '   A is exchange where at least one payment is on account,
              '   CN is credit note to account,
              '   P is account payment
1900          If strExchType = "APP" Then
1910              CreateApproFromExchange strEXCHID, appTRID
                  'Consolidate lines if necessary
1920              ConsolidateApproLines appTRID, bMadeChange
1930          End If
1940          If strExchType = "AR" Then AUTOGenerateReturnAndInvoice strEXCHID 'takes lines inserted into tTmpBookclub and uses them
                                                                                'to create invoice and return
                                                                                'associates the invoice with an exchange
1950          If strType = "A" And strExchType <> "AR" Then   'This is a credit exchange
1960              CreateInvoiceFromExchange strEXCHID
1970          End If
1980          If strType = "P" Then   'This is a payment to an account
1990              CreatePaymentToAccount strEXCHID, lngLastPayment
2000          End If
2010          If strType = "CN" Then   'This is a payment to an account
2020              CreateCreditNoteFromExchange strEXCHID
2030          End If
              
2040      End Select
      '    If oPC.TransactionCount > 0 Then
      '        If bDebug Then LogSaveToFile "Committing: Number = " & CStr(pExchangeNumber)
      '      '  oPC.CommitTransaction
      '    Else
      '        LogSaveToFile "No transaction to commit:strType = " & strType & "   strExchType = " & strExchType & "  strErrPos " & strErrPos
      '    End If
          
SkipforZENDXEND:
2050      RaiseEvent DBStatusMsg("Waiting")
2060  strErrPos = "Pos 9"

2070      SaveSaleData = 0
2080      Exit Function
          
ErrHandlerControl:
2090      On Error Resume Next
2100      LogSaveToFile "ErrHandlerControl: ExchID = " & pExchID
2110      LogSaveToFile "ErrHandlerControl: ExchID = " & pExchID & Error
2120      If oPC.TransactionCount > 0 Then
2130          LogSaveToFile "Rolling back tansaction in SaveSaleData"
2140          oPC.RollbackTransaction
2150      End If
2160      SaveSaleData = 100 + Result
2170      Exit Function
errHandler:

2180      ErrPreserve
2190      LogSaveToFile "ErrPreserve: ExchID = " & pExchID
2200      SaveSaleData = 200 + Result
2210      If ErrMustStop Then Debug.Assert False: Resume
2220      ErrorIn "Z_PollingServices.SaveSaleData(pSalesMessage)", Array(pSalesMessage), , oPC.COShort, "strErrPos", Array(strErrPos)
End Function


Private Function ProcessPayments(pPaymentPart As String) As Integer
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
         lngLastPayment = IIf(FNS(strExchType) = "PCC", lngAmt * -1, lngAmt)
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
         cmd.CommandTimeout = 180
         cmd.Execute
         lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
         If lngErrstatus <> 0 Then
             LogSaveToFile "Error in Stored procedure sp_UPDATEPOS_2, Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@ERRPOSITION"))
             GoTo ErrHandlerControl
         End If
    
         If strPaymentType = "RD" Then  'Deposit is redeemed - must mark C.O.L deposit status to 'RD' (redeemed)]
             strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'E' WHERE COL_ID = " & lngPaymentCOLID
             oPC.COShort.Execute strUpdate
         ElseIf strPaymentType = "RDEP" Then 'Deposit is refunded
             strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'X' WHERE COL_ID = " & lngPaymentCOLID
             oPC.COShort.Execute strUpdate
         End If
         Set cmd = Nothing
    Next i
    ProcessPayments = 0
    Exit Function

ErrHandlerControl:
   ' oPC.COShort.RollbackTrans
    ProcessPayments = 1
    Exit Function
    
errHandler:
    ErrPreserve
    ProcessPayments = 2
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ProcessPayments(pPaymentPart)", pPaymentPart
End Function
Public Function ProcessCSLs(pCSLPart As String, pDate As String) As Integer
10        On Error GoTo errHandler
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
        Dim lngSupervisor As Long
20        tmp = Split(pCSLPart, "~")
30        For i = 0 To UBound(tmp)
40            lin = Split(tmp(i), vbTab)
              
50            strPID = FNS(lin(0))
60            lngCOLID = FNN(lin(1))
70            lngQty = FNN(lin(2))
80            lngPrice = FNN(lin(3))
90            lngPriceAlteration = FNN(lin(4))
100           lngDiscount = FNN(lin(5))
110           dblDiscountRate = FNDBL(lin(6))
120           dblVATRATE = FNDBL(lin(7))
130           strCounterfoil = FNS(lin(8))
140           strDiscountDescription = FNS(lin(9))
            If UBound(lin) > 9 Then
              lngSupervisor = FNN(lin(10))
            Else
                lngSupervisor = 0
            End If
          ''
150           Set cmd = Nothing
160           Set cmd = New ADODB.Command
170           cmd.ActiveConnection = oPC.COShort
              
180           cmd.CommandType = adCmdStoredProc
190           cmd.NamedParameters = True
200           If strExchType = "APP" Then
210               cmd.CommandText = "sp_APPLine"
220               Set prm = cmd.CreateParameter("@APPLID", adInteger, adParamInput, , lngCOLID)
230               cmd.Parameters.Append prm
240               Set prm = Nothing
250           ElseIf strExchType = "AR" Then
260               cmd.CommandText = "sp_APPRET"   'Just puts data into tTmpBookclub table for processing by sp_APPRO_INV_AND_RETURN_POS
270               Set prm = cmd.CreateParameter("@APPLID", adInteger, adParamInput, , lngCOLID)
280               cmd.Parameters.Append prm
290               Set prm = Nothing
300           ElseIf strExchType <> "DEP" And strExchType <> "RDEP" And strExchType <> "OR" Then
310               cmd.CommandText = "sp_CSL"
320           Else
330               If strExchType = "DEP" Then  'Deposit is accepted for order
340                   strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'P' WHERE COL_ID = " & lngCOLID
350                   oPC.COShort.Execute strUpdate
360               End If
370               If strExchType = "RDEP" Then  'Deposit is accepted for order
380                     strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'X',COL_FUlfilled = 'CAN' WHERE COL_ID = " & lngCOLID
390                     oPC.COShort.Execute strUpdate
                        strUpdate = "UPDATE tProduct SET P_QTYONBACKORDER = ISNULL(P_QTYONBACKORDER,0) - (ISNULL(COL_QTY,0) - ISNULL(COL_QTYDISPATCHED,0)) FROM tCOL JOIN tPRODUCT ON COL_P_ID = P_ID WHERE COL_ID = " & lngCOLID
                        oPC.COShort.Execute strUpdate
                        strUpdate = "UPDATE tTR SET TR_STATUS = 4 WHERE TR_ID IN " _
                        & " (SELECT TR_ID FROM tCOL JOIN tTR ON tCOL.COL_TR_ID = tTR.TR_ID WHERE (TR_STATUS =3 AND COL_DateReplaced IS NULL and COL_ID = " & CStr(lngCOLID) & ")" _
                        & " GROUP BY TR_CODE ,TR_ID HAving (MAX(COL_Fulfilled) = 'CAN') or (MAX(COL_Fulfilled) = 'FUL'))"
                      oPC.COShort.Execute strUpdate
400               End If
410               Set cmd = Nothing
420               Exit Function
430           End If
       '   MsgBox "ProcessCSLs:2 " & CStr(oPC.TransactionCount) & " " & strExchType & " " & cmd.Name
440           Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
450           cmd.Parameters.Append prm
460           Set prm = Nothing
          
470           Set prm = cmd.CreateParameter("@PID", adVarChar, adParamInput, 40, strPID)
480           cmd.Parameters.Append prm
490           Set prm = Nothing
          
500           Set prm = cmd.CreateParameter("@Qty", adInteger, adParamInput, , lngQty)
510           cmd.Parameters.Append prm
520           Set prm = Nothing
          ''
530           Set prm = cmd.CreateParameter("@Price", adInteger, adParamInput, , lngPrice)
540           cmd.Parameters.Append prm
550           Set prm = Nothing
560           Set prm = cmd.CreateParameter("@PriceAlteration", adInteger, adParamInput, , lngPriceAlteration)
570           cmd.Parameters.Append prm
580           Set prm = Nothing
             ' MsgBox "Problem with price alteration"
590           Set prm = cmd.CreateParameter("@DiscountRate", adNumeric, adParamInput, , dblDiscountRate)
600           prm.NumericScale = 2
610           prm.Precision = 8
620           cmd.Parameters.Append prm
630           Set prm = Nothing
          ''
640           Set prm = cmd.CreateParameter("@Discount", adInteger, adParamInput, , lngDiscount)
650           cmd.Parameters.Append prm
660           Set prm = Nothing
          
670           Set prm = cmd.CreateParameter("@VRate", adInteger, adParamInput, , dblVATRATE)
680           cmd.Parameters.Append prm
690           Set prm = Nothing
          
700           Set prm = cmd.CreateParameter("@STID", adInteger, adParamInput, , oPC.DefaultStoreID)
710           cmd.Parameters.Append prm
720           Set prm = Nothing
          
730           Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , IIf(lngTPID = 0, oPC.CustomerTPID, lngTPID))
740           cmd.Parameters.Append prm
750           Set prm = Nothing
          ''
760           Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , strEXCHID)
770           cmd.Parameters.Append prm
780           Set prm = Nothing
          '
790           Set prm = cmd.CreateParameter("@COUNTERFOIL", adVarChar, adParamInput, 30, strCounterfoil)
800           cmd.Parameters.Append prm
810           Set prm = Nothing
          
820           Set prm = cmd.CreateParameter("@DISCOUNTDESCRIPTION", adVarChar, adParamInput, 100, strDiscountDescription)
830           cmd.Parameters.Append prm
840           Set prm = Nothing
          
850           Set prm = cmd.CreateParameter("@ID", adInteger, adParamOutput) 'dummy in this context
860           cmd.Parameters.Append prm
870           Set prm = Nothing
          
880           Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
890           cmd.Parameters.Append prm
900           Set prm = Nothing
          
910           Set prm = cmd.CreateParameter("@pPOSITION", adInteger, adParamOutput)
920           cmd.Parameters.Append prm
930           Set prm = Nothing
              
940           Set prm = cmd.CreateParameter("@ILID", adInteger, adParamInput, , lngCOLID)
950           cmd.Parameters.Append prm
960           Set prm = Nothing
              
970           Set prm = cmd.CreateParameter("@Dte", adVarChar, adParamInput, 35, pDate)
980           cmd.Parameters.Append prm
990           Set prm = Nothing
              
              Set prm = cmd.CreateParameter("@ActionSignature", adInteger, adParamInput, , lngSupervisor)
              cmd.Parameters.Append prm
              Set prm = Nothing
              
1000          cmd.CommandTimeout = 180
'Err.Raise 1444
1010          cmd.Execute
1020          lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
1030          If lngErrstatus <> 0 Then
1040              LogSaveToFile "Error in Stored procedure " & cmd.CommandText & " Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@pPOSITION"))
1050              GoTo ErrHandlerControl
1060          End If
                Set cmd = Nothing
1070      Next i
1080      ProcessCSLs = 0
1090      Exit Function
          
          
ErrHandlerControl:
      'MsgBox "ProcessCSLs:4 " & CStr(oPC.TransactionCount)
1100      ProcessCSLs = 1
1110      Exit Function

errHandler:
      'MsgBox "ProcessCSLs:5 " & CStr(oPC.TransactionCount)
1120      ErrPreserve

1130      ProcessCSLs = 2
1140      If ErrMustStop Then Debug.Assert False: Resume
1150      ErrorIn "Z_PollingServices.ProcessCSLs(pDate)", Array(FNS(pDate)), , , _
"TRID,strPID,lngQty,lngPrice,lngPriceAlteration,dblDiscountRate,lngDiscount,dblVATRATE,oPC.DefaultStoreID,TPID,strEXCHID,strCounterfoil,strDiscountDescription,lngCOLID,pDate", _
Array(TRID, strPID, lngQty, lngPrice, lngPriceAlteration, dblDiscountRate, lngDiscount, dblVATRATE, oPC.DefaultStoreID, IIf(lngTPID = 0, oPC.CustomerTPID, lngTPID), strEXCHID, strCounterfoil, strDiscountDescription, lngCOLID, pDate)
End Function
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
Public Sub CreateApproFromExchange(pExchID As String, appTRID As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_APPRO_FROMLIST"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , pExchID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@TRSTATUS", adInteger, adParamInput, , 3)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@STAFFID", adInteger, adParamInput, , 0)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@AID", adInteger, adParamOutput)
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
    
    If cmd.Parameters("@POSITION") <> 0 Then
        MsgBox "Error in CreateApproFromExchange"
    End If
    appTRID = FNN(cmd.Parameters("@AID"))
    Set cmd = Nothing

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.CreateApproFromExchange(pExchID,appTRID)", Array(pExchID, appTRID)
End Sub
Public Sub AUTOGenerateReturnAndInvoice(pExchID As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateInvoiceAndReturn_FromPOSAppro"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , pExchID)
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
    
    If cmd.Parameters("@POSITION") <> 0 Then
        MsgBox "Error in AUTOGenerateReturnAndInvoice"
    End If
    Set cmd = Nothing



    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.AUTOGenerateReturnAndInvoice(pEXCHID)", pExchID
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

Private Sub CreateInvoiceFromExchange(pExchID As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateInvoiceFromExchange"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , pExchID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@TRSTATUS", adInteger, adParamInput, , 4)  'Create the invoice as issued
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    
    cmd.Execute
    
    Set cmd = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.CreateInvoiceFromExchange(pEXCHID)", pExchID
End Sub

Private Sub CreateCreditNoteFromExchange(pExchID As String)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateCreditNoteFromExchange"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , pExchID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@TRSTATUS", adInteger, adParamInput, , 3)  'Create the invoice as issued
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    
    cmd.Execute
    
    Set cmd = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.CreateCreditNoteFromExchange(pEXCHID)", pExchID
End Sub
Private Sub CreatePaymentToAccount(pExchID As String, pAmt As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreatePaymentToAccount"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , pExchID)
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@TRSTATUS", adInteger, adParamInput, , 3)  'Create the Payment Transaction
    cmd.Parameters.Append par
    Set par = Nothing
    Set par = cmd.CreateParameter("@Amount", adInteger, adParamInput, , pAmt)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    
    cmd.Execute
    
    Set cmd = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.CreatePaymentToAccount(pEXCHID)", pExchID
End Sub

Private Sub CheckProductUpdatesAvailable(T1 As Boolean, t2 As Boolean, t3 As Boolean, t4 As Boolean, t5 As Boolean, t6 As Boolean, t7 As Boolean, reslt As Long)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim iReturn As Long
    Set cmd = New ADODB.Command
    cmd.CommandText = "CheckProductUpdatesAvailable"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@1", adBoolean, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@2", adBoolean, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@3", adBoolean, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@4", adBoolean, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@5", adBoolean, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@6", adBoolean, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@7", adBoolean, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@Status", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    oPC.OpenDBSHort
    cmd.ActiveConnection = oPC.COShort  'THis is used only by PBKS_POSSVR so it must use this connection - DO NOT CHANGE
    cmd.CommandTimeout = 30   'Allow 40 seconds
    cmd.Execute
    T1 = FNB(cmd.Parameters(0))
    t2 = FNB(cmd.Parameters(1))
    t3 = FNB(cmd.Parameters(2))
    t4 = FNB(cmd.Parameters(3))
    t5 = FNB(cmd.Parameters(4))
    t6 = FNB(cmd.Parameters(5))
    t7 = FNB(cmd.Parameters(6))
    reslt = FNN(cmd.Parameters(7))
    Set cmd = Nothing
    Exit Sub
errHandler:
    ErrPreserve
    If Err = -2147217871 Then   'Timeout
        reslt = 2   'send status to caller and continue
        Err.Clear
        oPC.DisconnectDBShort
        Exit Sub
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.CheckProductUpdatesAvailable(t1,t2,t3,t4,t5,t6,t7)", Array(T1, t2, t3, _
         t4, t5, t6, t7)
End Sub



