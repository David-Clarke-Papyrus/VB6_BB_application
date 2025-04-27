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

Dim arQueuePath() As String
Dim strTillpoint As String

Public Event FileProcessed(Till As String) '0 = Other, 1 = Sale, 2 = Update
Public Event DataSent()
Public Event SwitchPolling(bOn As Boolean)
Dim bDBOK As Boolean

Private WithEvents POSEvent As MSMQEvent
Attribute POSEvent.VB_VarHelpID = -1
Private QPOS As MSMQQueue

Private WithEvents SVREvent As MSMQEvent
Attribute SVREvent.VB_VarHelpID = -1
Private QSVR As MSMQQueue

Dim QI As MSMQQueueInfo
Dim QSVRAck As MSMQQueue
Dim SVRMsg As MSMQMessage

Event ErrorRegistering(pStationName As String)
Event ClientStatusChange()
Public WithEvents qTimer As XTimer
Attribute qTimer.VB_VarHelpID = -1
Public Function PutMasterDataInOutbox() As Boolean
    On Error GoTo errHandler
Dim OpenResult As Integer
'Dim cnt As Long
'-------------------------------
    OpenResult = oPC.OpenDBShort
'-------------------------------
      oPC.COShort.Execute "SELECT TOP 2000 tProdUpdates.* into #PRODS FROM tProdUpdates"
      PlaceInQueue GetRecordset("SELECT * FROM #PRODS"), "PROD", "tProdUpdates"
      oPC.COShort.Execute "DROP table #PRODS"


      PlaceInQueue GetRecordset("SELECT * FROM tStaffMemberUpdate"), "STAF", "tStaffMemberUpdate"

      oPC.COShort.Execute "SELECT TOP 2000 tTPUpdate_Cust.* into #CUSTS FROM tTPUpdate_Cust"
      PlaceInQueue GetRecordset("SELECT * FROM #CUSTS"), "CUST", "tTPUpdate_Cust"
      oPC.COShort.Execute "DROP table #CUSTS"

      oPC.COShort.Execute "SELECT TOP 2000 tCOUPDATE.* into #ORDR FROM tCOUPDATE"
      PlaceInQueue GetRecordset("SELECT  * FROM #ORDR"), "ORDR", "tCOUpdate"
      oPC.COShort.Execute "DROP table #ORDR"

      PlaceInQueue GetRecordset("SELECT * FROM tMarketing_Changes"), "MARK", "tMarketing_Changes"

      oPC.COShort.Execute "SELECT TOP 2000 tAPPUpdate.* into #APPL FROM tAPPUpdate"
      PlaceInQueue GetRecordset("SELECT * FROM #APPL"), "APPL", "tAPPUpdate"
      oPC.COShort.Execute "DROP table #APPL"
          
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
'Public Sub ClearCustomersOnFD()
'    On Error GoTo errHandler
'Dim j As Integer
'
'        For j = 1 To UBound(arQueuePath)
'            QI.FormatName = "DIRECT=OS:" & arQueuePath(j)
'            Set QSVR = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
'            Set SVRMsg = New MSMQMessage
'            SVRMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
'            SVRMsg.Label = "ClearCustomers," & Format(Now, "dd/mm/yyyy HH:NN")
'            SVRMsg.Body = ""
'            SVRMsg.Send QSVR
'        Next j
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "Z_PollingServices.ClearCustomersOnFD"
'End Sub
Public Sub ClearOnFD(pInstruction As String)
    On Error GoTo errHandler
Dim j As Integer

        For j = 1 To UBound(arQueuePath)
            QI.FormatName = "DIRECT=OS:" & arQueuePath(j)
            Set QSVR = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
            Set SVRMsg = New MSMQMessage
            SVRMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
            SVRMsg.Label = pInstruction & "," & Format(Now, "dd/mm/yyyy HH:NN")
            SVRMsg.Body = ""
            SVRMsg.Send QSVR
        Next j
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.ClearCustomersOnFD"
End Sub

'Public Sub ClearProductsOnFD()
'    On Error GoTo errHandler
'Dim j As Integer
'
'        For j = 1 To UBound(arQueuePath)
'            QI.FormatName = "DIRECT=OS:" & arQueuePath(j)
'            Set QSVR = QI.Open(MQ_SEND_ACCESS, MQ_DENY_NONE)
'            Set SVRMsg = New MSMQMessage
'            SVRMsg.Delivery = MQMSG_DELIVERY_RECOVERABLE
'            SVRMsg.Label = "ClearProducts," & Format(Now, "dd/mm/yyyy HH:NN")
'            SVRMsg.Body = ""
'            SVRMsg.Send QSVR
'        Next j
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "Z_PollingServices.ClearProductsOnFD"
'End Sub
Private Sub Class_Initialize()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim i As Integer
Dim OpenResult As Integer
'
'-------------------------------
    OpenResult = oPC.OpenDBShort
'-------------------------------
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
    Set qTimer = New XTimer
    qTimer.Interval = 10000
    qTimer.Enabled = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.Class_Initialize", , EA_NORERAISE
    HandleError
End Sub


Sub qTimer_Tick()
    QPOS.Reset
    On Error Resume Next
    QPOS.EnableNotification POSEvent
End Sub

Private Sub PlaceInQueue(pRs As ADODB.Recordset, strLabel As String, strTableName As String)
    On Error GoTo errHandler
Dim j As Integer

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
Dim OpenResult As Integer

''''-------------------------------
'''    OpenResult = oPC.OpenDBShort
''''-------------------------------
        oPC.COShort.CommandTimeout = 0
        Select Case strTableName
        Case "tProdUpdates"
            oPC.COShort.Execute "DELETE tProdupdates FROM (SELECT * FROM #PRODS) as t1 WHERE tPRODUPDATES.PRU_GUID = t1.PRU_GUID"
        Case "tStaffMemberUpdate"
            oPC.COShort.Execute "DELETE tStaffMemberUpdate"
        Case "tTPUpdate_Cust"
            oPC.COShort.Execute "DELETE tTPUpdate_Cust FROM (SELECT * FROM #CUSTS) as t1 WHERE tTPUpdate_Cust.CU_GUID = t1.CU_GUID"
        Case "tCOUpdate"
            oPC.COShort.Execute "DELETE tCOUpdate FROM (SELECT * FROM #ORDR) as t1 WHERE tCOUpdate.COU_GUID = t1.COU_GUID"
        Case "tMarketing_Changes"
            oPC.COShort.Execute "DELETE tMarketing_Changes"
        Case "tAPPUpdate"
            oPC.COShort.Execute "DELETE tAPPUpdate FROM (SELECT * FROM #APPL) as t1 WHERE tAPPUpdate.APPL_GUID = t1.APPL_GUID"
        End Select
            RaiseEvent DataSent
    End If
    If pRs.State = adStateOpen Then pRs.Close
''''---------------------------------------------------
'''    If OpenResult = 0 Then oPC.DisconnectDBShort
''''---------------------------------------------------
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

    On Error GoTo errHandler

    Set QPOS = Queue
    

    Set POSmsg = QPOS.Receive()
    QPOS.EnableNotification POSEvent
    If Not (POSmsg Is Nothing) Then
    
        If POSmsg.Label = "STOPPOLLING" Then
            RaiseEvent SwitchPolling(False)
            GoTo EXIT_Handler
        ElseIf POSmsg.Label = "STARTPOLLING" Then
            RaiseEvent SwitchPolling(True)
            GoTo EXIT_Handler
        End If
        If Left(POSmsg.Label, 4) = "POSN" Or Left(POSmsg.Label, 4) = "XEND" Or Left(POSmsg.Label, 4) = "ZEND" Then
            Set oStream = New ADODB.Stream
            oStream.Open
            oStream.WriteText POSmsg.Body
            oStream.Position = 0
            rs.Open oStream
        Else
            Set rs = POSmsg.Body
        End If
        
        SaveSaleData rs, lngResult
        RaiseEvent FileProcessed(strTillpoint)
        If lngResult = 0 Then    'Update sales data is OK
            rs.Close
        ElseIf lngResult = 1 Then
            rs.Close
        ElseIf lngResult = 2 Then   'Could not update successfully: preserve file
            On Error Resume Next
            rs.Save POSmsg.Label & ".bak", adPersistADTG
            On Error GoTo errHandler
        End If
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
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.POSEvent_Arrived"
End Sub


'~~~~~~~~~ SaveSaleData ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
'   Any data checking and manipulation prior to saving could be added to this
'   function, e.g. checking and creating valid product codes, checking of sale
'   prices etc.
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function SaveSaleData(rs As ADODB.Recordset, pResult As Integer)
    On Error GoTo errHandler
Dim rsOP As ADODB.Recordset
Dim rsEXCH As ADODB.Recordset
Dim rsCSL As ADODB.Recordset
Dim rsPAY As ADODB.Recordset
Dim msg As String
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim TRID As Long
Dim strUpdate As String
Dim lngErrstatus As Long
Dim bOK As Integer
Dim bMadeChange As Boolean

strTillpoint = ""
    bOK = 99
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBShort
'-------------------------------
    oPC.COShort.BeginTrans
    'Update or add to tZSession
    'Update or add to tCS
    'Update or add to tOPSession
    'Add to tExchange
    Select Case rs.Fields("typ")
    Case "O"
        strSQL = "UPDATE tCS SET CS_ENDSession = '" & ReverseDateTime(rs!OPS_ENDTIME) & "' WHERE CS_GUID = '" & rs!OPS_ID & "'"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandText = strSQL
        cmd.CommandType = adCmdText
        cmd.Execute
    Case "Z"
        strSQL = "UPDATE tZSession SET Z_ENDDATE = '" & ReverseDateTime(rs!Z_ENDDATE) & "' WHERE Z_ID = '" & rs!Z_ID & "'"
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = oPC.COShort
        cmd.CommandText = strSQL
        cmd.CommandType = adCmdText
        cmd.Execute
'        RaiseEvent ZSent
    Case "E"
            strTillpoint = FNS(rs!Z_Tillpoint)
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = oPC.COShort
            cmd.CommandText = "sp_UPDATEPOS_1"
            cmd.CommandType = adCmdStoredProc
            cmd.NamedParameters = True
            
            Set prm = cmd.CreateParameter("@ZID", adGUID, adParamInput, , rs!Z_ID)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@STARTDATE", adDate, adParamInput, , rs!Z_STARTDATE)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@ENDDATE", adDate, adParamInput, , rs!Z_ENDDATE)
            cmd.Parameters.Append prm
            Set prm = Nothing
                        
            Set prm = cmd.CreateParameter("@NOMINALDATE", adDate, adParamInput, , rs!Z_NOMINALDATE)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@TILLPOINT", adVarChar, adParamInput, 20, rs!Z_Tillpoint)
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set rsOP = rs.Fields("OPSession").Value
            
            Set prm = cmd.CreateParameter("@OPSID", adGUID, adParamInput, , rsOP!OPS_ID)
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set prm = cmd.CreateParameter("@OPOPERATORID", adInteger, adParamInput, , rsOP!OPS_OPERATORID)
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set prm = cmd.CreateParameter("@OPSSTARTTIME", adDate, adParamInput, , rsOP!OPS_STARTTIME)
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set prm = cmd.CreateParameter("@OPSENDTIME", adDate, adParamInput, , rsOP!OPS_ENDTIME)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
                   
            Set rsEXCH = rsOP.Fields("POSEXCHANGE").Value
            
            Set prm = cmd.CreateParameter("@EXCHOPERATORID", adInteger, adParamInput, , rsEXCH!EXCH_OPERATORID)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@EXCHSUPERVISORID", adInteger, adParamInput, , rsEXCH!EXCH_SUPERVISORID)
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@EXCHNUMBER", adInteger, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_NUMBER))
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_ID))
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set prm = cmd.CreateParameter("@EXZSESSID", adGUID, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_ZSESSIONID))
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set prm = cmd.CreateParameter("@EXOPSESSID", adGUID, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_OPSESSIONID))
            cmd.Parameters.Append prm
            Set prm = Nothing
            
            Set prm = cmd.CreateParameter("@EXCHDATE", adDate, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_SALEDATE))
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set prm = cmd.CreateParameter("@EXCHSALESVALUE", adInteger, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_SALEVALUE))
            cmd.Parameters.Append prm
            Set prm = Nothing
            Set prm = cmd.CreateParameter("@EXCHDISCOUNTVALUE", adInteger, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_DISCOUNTVALUE))
            cmd.Parameters.Append prm
            Set prm = Nothing
            Set prm = cmd.CreateParameter("@EXCHVATVALUE", adInteger, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_VATVALUE))
            cmd.Parameters.Append prm
            Set prm = Nothing
            Set prm = cmd.CreateParameter("@EXCHCHANGEGIVEN", adInteger, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_CHANGEGIVEN))
            cmd.Parameters.Append prm
            Set prm = Nothing
            Set prm = cmd.CreateParameter("@EXCHLOYALTYVALUE", adInteger, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_LOYALTYVALUE))
            cmd.Parameters.Append prm
            Set prm = Nothing
            Set prm = cmd.CreateParameter("@EXCHTYPE", adVarChar, adParamInput, 10, IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_TYPE))
            cmd.Parameters.Append prm
            Set prm = Nothing
            'This willalmost always come thru to the server before it has been voided
        '    Set prm = cmd.CreateParameter("@EXCHVOIDED", adBoolean, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_VOIDED))
        '    cmd.Parameters.Append prm
        '    Set prm = Nothing
            Set prm = cmd.CreateParameter("@EXCHVOIDS", adInteger, adParamInput, , IIf(rsEXCH.EOF, Null, rsEXCH!EXCH_VOIDS))
            cmd.Parameters.Append prm
            Set prm = Nothing
                  
            Set prm = cmd.CreateParameter("@EXCHCUSTOMERNAME", adVarChar, adParamInput, 100, IIf(rsEXCH.EOF, Null, Left(rsEXCH!EXCH_NOTE, 100)))
            cmd.Parameters.Append prm
            Set prm = Nothing
           
            Set prm = cmd.CreateParameter("@EXCHTPID", adInteger, adParamInput, , IIf(FNN(rsEXCH!EXCH_TP_ID) = 0, oPC.CustomerTPID, rsEXCH!EXCH_TP_ID))
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
           
            On Error Resume Next
            cmd.Execute
            lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
            If lngErrstatus = 51 Then
                GoTo ErrHandlerControl2
            End If
            On Error GoTo errHandler
            
            If lngErrstatus <> 0 Then
                LogSaveToFile "Error in Stored procedure sp_UPDATEPOS_1 Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@ERRPOSITION"))
                GoTo ErrHandlerControl
            End If
            
            TRID = IIf(IsNull(cmd.Parameters("@TRID")), 0, cmd.Parameters("@TRID"))
            
        'Add  to tCSL
            Set rsCSL = rsEXCH.Fields("SALESLINES").Value
            oPC.COShort.Execute "DELETE FROM tTMPBOOKCLUB"
            Do While Not rsCSL.EOF
                Set cmd = Nothing
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = oPC.COShort
                cmd.CommandType = adCmdStoredProc
                cmd.NamedParameters = True
                If rsEXCH!EXCH_TYPE = "APP" Then
                    cmd.CommandText = "sp_APP"
                    
                ElseIf rsEXCH!EXCH_TYPE = "AR" Then
                    cmd.CommandText = "sp_APPRET"
                    Set prm = cmd.CreateParameter("@APPLID", adInteger, adParamInput, , FNN(rsCSL!CSL_COLID))
                    cmd.Parameters.Append prm
                    Set prm = Nothing
                ElseIf rsEXCH!EXCH_TYPE <> "DEP" And rsEXCH!EXCH_TYPE <> "RDEP" And rsEXCH!EXCH_TYPE <> "OR" Then
                    cmd.CommandText = "sp_CSL"
                Else
                    If FNS(rsEXCH!EXCH_TYPE) = "DEP" Then  'Deposit is accepted for order
                        strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'P' WHERE COL_ID = " & FNN(rsCSL!CSL_COLID)
                        oPC.COShort.Execute strUpdate
                    End If
                    If FNS(rsEXCH!EXCH_TYPE) = "RDEP" Then  'Deposit is accepted for order
                        strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'X' WHERE COL_ID = " & FNN(rsCSL!CSL_COLID)
                        oPC.COShort.Execute strUpdate
                    End If
                    GoTo SKIPCSL
                End If
                
                Set prm = cmd.CreateParameter("@TRID", adInteger, adParamInput, , TRID)
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@PID", adGUID, adParamInput, , rsCSL.Fields("CSL_P_ID"))
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@Qty", adInteger, adParamInput, , rsCSL.Fields("CSL_QTY"))
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@Price", adInteger, adParamInput, , rsCSL.Fields("CSL_PRICE"))
                cmd.Parameters.Append prm
                Set prm = Nothing
                Set prm = cmd.CreateParameter("@PriceAlteration", adInteger, adParamInput, , rsCSL.Fields("CSL_PRICEALTERATION"))
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@DiscountRate", adNumeric, adParamInput, , rsCSL.Fields("CSL_DISCOUNTRate"))
                prm.NumericScale = 2
                prm.Precision = 8
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@Discount", adInteger, adParamInput, , rsCSL.Fields("CSL_Discount"))
                cmd.Parameters.Append prm
                Set prm = Nothing
               
                Set prm = cmd.CreateParameter("@VRate", adInteger, adParamInput, , rsCSL.Fields("CSL_VATRATE"))
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@STID", adInteger, adParamInput, , oPC.DefaultStoreID)
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , IIf(IsNull(rsEXCH!EXCH_TP_ID), oPC.CustomerTPID, rsEXCH!EXCH_TP_ID))
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , rsEXCH!EXCH_ID)
                cmd.Parameters.Append prm
                Set prm = Nothing
    '
                Set prm = cmd.CreateParameter("@COUNTERFOIL", adVarChar, adParamInput, 30, FNS(rsCSL.Fields("CSL_Counterfoil")))
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@DISCOUNTDESCRIPTION", adVarChar, adParamInput, 100, FNS(rsCSL.Fields("CSL_DiscountDescription")))
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
                
            
            
              '  On Error Resume Next
                cmd.Execute
                lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
                On Error GoTo errHandler
                
                If lngErrstatus <> 0 Then
                    LogSaveToFile "Error in Stored procedure " & cmd.CommandText & " Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@pPOSITION"))
                    GoTo ErrHandlerControl
                End If
                
                
                rsCSL.MoveNext
                       
            Loop
            If rsEXCH!EXCH_TYPE = "APP" Then   'Consolidate lines if necessary
                ConsolidateApproLines TRID, bMadeChange
            End If
SKIPCSL:
        If rsEXCH!EXCH_TYPE <> "APP" Then

        'Add to tPayment
            Set rsPAY = rsEXCH.Fields("PAYMENTS").Value
            Do While Not rsPAY.EOF
                Set cmd = Nothing
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = oPC.COShort
                cmd.CommandType = adCmdStoredProc
                cmd.NamedParameters = True
                cmd.CommandText = "sp_UPDATEPOS_2"
                Set prm = cmd.CreateParameter("@EXCHID", adGUID, adParamInput, , rsEXCH!EXCH_ID)
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@PAYTYPE", adChar, adParamInput, 4, rsPAY!PAY_PAYMENTTYPE)
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@PAYAMT", adInteger, adParamInput, , IIf(FNS(rsEXCH!EXCH_TYPE) = "PCC", FNN(rsPAY!PAY_AMT) * -1, FNN(rsPAY!PAY_AMT)))
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@PAYREF", adVarChar, adParamInput, 100, rsPAY!PAY_REF)
                cmd.Parameters.Append prm
                Set prm = Nothing
                Set prm = cmd.CreateParameter("@PAYNOTE", adVarChar, adParamInput, 50, rsPAY!PAY_NOTE)
                cmd.Parameters.Append prm
                Set prm = Nothing
               
                
                Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput)
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                Set prm = cmd.CreateParameter("@ErrPOSITION", adInteger, adParamOutput)
                cmd.Parameters.Append prm
                Set prm = Nothing
                
                On Error Resume Next
                cmd.Execute
                lngErrstatus = FNN(cmd.Parameters("@pErrCode"))
                On Error GoTo errHandler
                
                If lngErrstatus <> 0 Then
                    LogSaveToFile "Error in Stored procedure sp_UPDATEPOS_2, Error code is " & CStr(lngErrstatus) & " in position: " & CStr(cmd.Parameters("@ERRPOSITION"))
                    GoTo ErrHandlerControl
                End If
                
                
                
                If rsPAY!PAY_PAYMENTTYPE = "RD" Then  'Deposit is redeemed - must mark C.O.L deposit status to 'RD' (redeemed)]
                    strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'E' WHERE COL_ID = " & FNN(rsPAY!PAY_COLID)
                    oPC.COShort.Execute strUpdate
                ElseIf FNS(rsEXCH!EXCH_TYPE) = "RDEP" Then 'Deposit is refunded
                    strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'X' WHERE COL_ID = " & FNN(rsPAY!PAY_COLID)
                    oPC.COShort.Execute strUpdate
'                ElseIf FNS(rsEXCH!EXCH_TYPE) = "DEP" Then  'Deposit is accepted for order
'                    strUpdate = "UPDATE tCOL SET COL_DepositStatus = 'P' WHERE COL_ID = " & FNN(rsCSL!CSL_COLID)
'                    oPC.CO.Execute strUpdate
                End If
                
                rsPAY.MoveNext
            Loop
            End If
            'If this exchange is set to void another then we must do so
            If IIf(rsEXCH.EOF, 0, rsEXCH!EXCH_VOIDS) > 0 Then
                VoidAnExchange rs!Z_ID, rsEXCH!EXCH_VOIDS
            End If
            Set cmd = Nothing
        If rsEXCH!EXCH_TYPE = "AR" Then
            AUTOGenerateReturnAndInvoice rsEXCH!EXCH_ID   'takes lines inserted into tTmpBookclub and uses them to create invoice and return
        End If                                                        'associates the invoice with and exchange
    End Select
    
    bOK = True
    
    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    pResult = 0
    Exit Function
    
ErrHandlerControl:
    oPC.COShort.RollbackTrans
    pResult = 2
    Exit Function
ErrHandlerControl2:
    oPC.COShort.RollbackTrans
    pResult = 1
    Exit Function
errHandler:
    ErrPreserve
    pResult = 99
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "Z_PollingServices.SaveSaleData(rs)", rs, EA_ROLLBACK + EA_NORERAISE + EA_ADVANCED, oPC.COShort
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

