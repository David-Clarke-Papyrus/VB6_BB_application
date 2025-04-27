VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Batch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim fErr
Dim fMode
Dim strMMsg As String

Public Enum PreviewPrint
    PrintReport = 1
    PreviewReport = 2
End Enum
Public Event Progress(lngPos As Long, lngMax As Long)
Public Event ProgressB(lngPos As Long, lngMax As Long, msg As String)
Public Event STATUS(msg As String)
Public Event OrderStatus(RequestStatus As String, Answer As Boolean)
Public Event DEResult(pMsg As String, pErr As Long)
'Dim WithEvents oCD As Z_CDWrite
Dim mNominalDate As Date
Dim gSTAFFID As Long

Public Sub RunPapyrusSchedulerDayend()
    On Error GoTo errHandler

Dim cmdADO As ADODB.Command
Dim prm As ADODB.Parameter
Dim pResult As String

Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmdADO = New ADODB.Command
    cmdADO.ActiveConnection = oPC.COShort
    cmdADO.CommandText = "RunDayendManually"
    cmdADO.CommandType = adCmdStoredProc
    Set prm = Nothing
    Set prm = New ADODB.Parameter
    prm.Type = adDate
    prm.Direction = adParamInput
    prm.Value = Date
    cmdADO.Parameters.Append prm
    Set prm = Nothing
    
    cmdADO.execute
    oPC.CloseScheduler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.RunPapyrusSchedulerDayend"
End Sub
Public Sub SetMainMsg(pMsg As String)
    On Error GoTo errHandler
    strMMsg = pMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.SetMainMsg(pMsg)", pMsg
End Sub
Function DeleteTableIfExists(pTableName As String) As Boolean
    On Error GoTo errHandler
Dim cmdADO As ADODB.Command
Dim prm As ADODB.Parameter
Dim pResult As String

    Set cmdADO = New ADODB.Command
    cmdADO.ActiveConnection = oPC.CO
    cmdADO.CommandText = "DropTableIfExists"
    cmdADO.CommandType = adCmdStoredProc
    Set prm = Nothing
    Set prm = New ADODB.Parameter
    prm.Type = adVarChar
    prm.Direction = adParamInput
    prm.Size = 50
    prm.Value = pTableName
    cmdADO.Parameters.Append prm
    Set prm = Nothing
    Set prm = New ADODB.Parameter
    prm.Type = adVarChar
    prm.Direction = adParamInput
    prm.Size = 5
    prm.Value = pResult
    cmdADO.Parameters.Append prm
    
    cmdADO.execute
    
    DeleteTableIfExists = IIf(UCase(cmdADO.Parameters(1)) = "TRUE", True, False)
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.DeleteTableIfExists(pTableName)", pTableName
End Function

'
Public Function DailyUpdate(pNominalDate As Date, gSTAFFID As Long) As Long
    On Error GoTo errHandler

Dim iWeekToRemove, iMonthToRemove, iCurrentWeek, iCurrentMonth As Integer
Dim strConn As String
Dim cmd As ADODB.Command
Dim Param As ADODB.Parameter
Dim lngResult As Long
Dim lngPosition As Long
Dim lngOPID As Long
Dim ps As z_papysystem
Dim strError As String
Dim strERR As String
Dim X As String, Y As String
Dim f As File
Dim strSQL As String
Dim tmpPapyConn As PapyConn
Dim dteLastCopied As Date
Dim strParentFolderName As String
Dim retval
Dim oError As a_Error
Dim rs As ADODB.Recordset
Dim prm As ADODB.Parameter
Dim oOP As a_Operation
Dim strPosition As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    RaiseEvent STATUS("Dayend running . . .")
    DoEvents
        
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "sp_DAYEND"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@NOMINALDATE", adDate, adParamInput, , pNominalDate)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@SMID", adInteger, adParamInput, , gSTAFFID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@pErrCode", adInteger, adParamOutput, , 0)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@pPosition", adInteger, adParamOutput, , 0)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@OPID", adInteger, adParamOutput, , 0)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    cmd.execute
    If IIf(IsNull(cmd.Parameters(2)), 0, FNN(cmd.Parameters(2))) <> 0 Or IIf(IsNull(cmd.Parameters(3)), 0, FNN(cmd.Parameters(3))) <> 66 Then
        strPosition = CStr(cmd.Parameters(3))
        strError = CStr(cmd.Parameters(2))
    End If
'
    Set ps = New z_papysystem
    Set oOP = New a_Operation
    If Not IsNull(cmd.Parameters(4)) Then
        oOP.Load lngResult, CLng(cmd.Parameters(4))
        oOP.BeginEdit
        If strError > "" Then
            oOP.Fullreport = "Error: " & strError & vbCrLf & "at position: " & strPosition & vbCrLf & oOP.Fullreport & vbCrLf & ps.GetDiskStatus()
        Else
            oOP.Fullreport = oOP.Fullreport & ps.GetDiskStatus()
        End If
        oOP.ApplyEdit
    End If
    Set oOP = Nothing
    Set ps = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    If strError > "" Then GoTo SP_ERR_Handler
 '  oPC.CO.Close
EXIT_Handler:
    DailyUpdate = 0
    DoEvents
    Set cmd = Nothing
    Set ps = Nothing
    Exit Function

SP_ERR_Handler:
    DailyUpdate = 99
    RaiseEvent DEResult("The dayend has failed at position " & strPosition & " with error code " & strError & ". Contact Papyrus support", 99)
    DoEvents
    Set cmd = Nothing
    Set ps = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.DailyUpdate(pNominalDate)", pNominalDate
End Function
Public Function SetOOS()
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.SetOOS"
End Function
Public Function RemoveSOs(pDate As Date, pContactID As Long)
    On Error GoTo errHandler

Dim oError As a_Error
Dim strConn As String
Dim objOp As a_Operation
Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    Set objOp = New a_Operation
    objOp.BeginEdit
        fMode = "ADD"
        objOp.StartedAt = Now()
        objOp.TypeID = 3
        objOp.NominalDate = pDate
        objOp.OperatorID = pContactID
    objOp.ApplyEdit lngResult
    cmd.CommandText = "RemoveOLDSOs_1"
    cmd.CommandType = adCmdStoredProc
    cmd.execute , Array(pDate)
    retval = RunProc("RemoveOLDSOs_2", "", "", oPC.CO)
   ' retval = RunProc("RemoveOLDSOs_3", oPC.CO, "") 'cannot run this as it deletes the links between product and matching tables
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.RemoveSOs(pDate,pContactID)", Array(pDate, pContactID)
End Function


Function RebuildPublisherTable()
    On Error GoTo errHandler
Dim strSQL As String

    strSQL = "SELECT  P_Publisher INTO PublisherLIst_TEMP FROM tPRODUCT GROUP BY p_Publisher"
    retval = Me.DropTable("PublisherLIst_TEMP", "Erasing Table  . . . ")
    
    Me.RunSQL strSQL, oPC.CO '"q_ListOfPublishers", Array(), "Building list of Publishers  . . . "
    
    SaveSetting "SupplierChange", "Settings", "LastPubListRebuild", Now

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.RebuildPublisherTable"
End Function

Function SupplierChange(pcolPublisher As Collection, pLastTPID, pNewTPID, pMsg As String, Optional pConn As ADODB.Connection) As Long
    On Error GoTo errHandler


Dim cmd As ADODB.Command
Dim lngRecordsAffected As Long
Dim strSQL As String
Dim strWHERE As String
Dim i As Integer

    Set cmd = New ADODB.Command
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.CO
    Else
        cmd.ActiveConnection = pConn
    End If
    
    strWHERE = ""
    For i = 1 To pcolPublisher.Count
        If Len(strWHERE) > 0 Then
            strWHERE = strWHERE & " OR "
        End If
        strWHERE = strWHERE & " tProduct.P_Publisher = '" & Replace(pcolPublisher(i), "'", "''") & "'"
    Next i
    If pLastTPID > 0 Then
    strSQL = "UPDATE tProduct SET tProduct.P_SupplierID = " & pNewTPID _
                    & " WHERE (( " & strWHERE & " ) " _
                    & " AND ((tProduct.P_SupplierID)= " & pLastTPID & ")); "
    Else
    strSQL = "UPDATE tProduct SET tProduct.P_SupplierID = " & pNewTPID _
                    & " WHERE (( " & strWHERE & " ) " _
                    & " AND ((tProduct.P_SupplierID)= " & pLastTPID & ")); "
    End If
    Me.RunSQL strSQL
    
EXIT_Handler:
    Set cmd = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.SupplierChange(pcolPublisher,pLastTPID,pNewTPID,pMsg,pConn)", _
         Array(pcolPublisher, pLastTPID, pNewTPID, pMsg, pConn)
End Function

Public Sub LoadStockSummary()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim dteDate As Date
Dim dteDate2 As Date
Dim strSQL As String
Dim RetailValue As Currency
Dim CostValue As Currency
Dim lngWeekNum As Long
Dim lngRecordsReturned As Long

'    Set rs = New ADODB.Recordset
'    strMMsg = "Calculating Stock Summary"
''    objOp.NominalDate = #4/3/2001#
'    dteDate2 = Format(objOp.NominalDate, "dd/mm/yyyy")
'    dteDate = Format(DateAdd("d", -1, dteDate2), "dd/mm/yyyy")
'    lngWeekNum = oPC.Configuration.CurrentWeek2(dteDate2)
'
'    lngRecordsReturned = RunProc("q_GenerateStockSummary", Array(dteDate2, lngWeekNum), "Step 1")
'
''    **** Updating Cash Sales figures over period
'    lngRecordsReturned = RunGetRecordset("q_r13_21", enQuery, Array(dteDate, dteDate2), "Step 2", rs)
'    RetailValue = FNN(rs.Fields("Retailvalue"))
'    CostValue = FNN(rs.Fields("CostValue"))
'    If rs.State = adStateOpen Then rs.Close
'    Set rs = Nothing
'    lngRecordsReturned = RunProc("q_r13_21_1", Array(RetailValue, CostValue, lngWeekNum, dteDate2), "Step 3")
'
''    **** Updating the Invoice sales figures over period
'    Set rs = New ADODB.Recordset
'    retval = RunGetRecordset("q_r13_22", enQuery, Array(dteDate, dteDate2), "Step 4", rs)
'    RetailValue = FNN(rs.Fields("RetailValue"))
'    CostValue = FNN(rs.Fields("CostValue"))
'    If rs.State = adStateOpen Then rs.Close
'    Set rs = Nothing
'    lngRecordsReturned = RunProc("q_r13_22_1", Array(RetailValue, CostValue, dteDate2), "Step 5")
'
''    **** Updating the deliveries in over period
'    Set rs = New ADODB.Recordset
'    retval = RunGetRecordset("q_r13_3", enQuery, Array(dteDate, dteDate2), "Step 6", rs)
'    RetailValue = FNN(rs.Fields("RetailValue"))
'    CostValue = FNN(rs.Fields("CostValue"))
'    If rs.State = adStateOpen Then rs.Close
'    Set rs = Nothing
'    lngRecordsReturned = RunProc("q_r13_3_1", Array(RetailValue, CostValue, dteDate2), "Step 7")
'
''    **** Updating the orders over period
'    Set rs = New ADODB.Recordset
'    retval = RunGetRecordset("q_r13_4", enQuery, Array(dteDate, dteDate2), "Step 8", rs)
'    RetailValue = FNN(rs.Fields("RetailValue"))
'    CostValue = FNN(rs.Fields("CostValue"))
'    If rs.State = adStateOpen Then rs.Close
'    Set rs = Nothing
'    lngRecordsReturned = RunProc("q_r13_4_1", Array(RetailValue, CostValue, dteDate2), "Step 9")
'
''    **** Updating the transferred in/out over period
'    Set rs = New ADODB.Recordset
'    retval = RunGetRecordset("q_r13_5", enQuery, Array(dteDate, dteDate2), "Step 10", rs)
'    RetailValue = (FNN(rs.Fields("RetailValue")) * -1)
'    CostValue = (FNN(rs.Fields("CostValue")) * -1)
'    If rs.State = adStateOpen Then rs.Close
'    Set rs = Nothing
'    lngRecordsReturned = RunProc("q_r13_5_1", Array(RetailValue, CostValue, dteDate2), "Step 11")
        
EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.LoadStockSummary"
End Sub

'Public Sub UpdateStockAgeFigures()
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim dteDate As Date
'Dim dteDate2 As Date
'Dim iBal As Integer
'Dim Months6 As Integer
'Dim Months12 As Integer
'Dim Months18 As Integer
'Dim Months18Plus As Integer
'
'    dteDate2 = Format(objOp.NominalDate, "dd/mm/yyyy")
''    dteDate2 = Format(#3/29/2001#, "dd/mm/yyyy")
'    'Obtains the value of stock on hand. Currently over the period of 18, 12 and 6 months
'    DropTable "TEMP_r13_6_1", ""
'    RunProc "q_r13_6_1", Array(), "Step 1"
'
''    ** Process received transactions
'    dteDate = DateAdd("m", -6, dteDate2)
'    RunProc "q_r13_6_11", Array(dteDate, dteDate2), "Step 2"
'    dteDate = DateAdd("m", -12, dteDate2)
'    RunProc "q_r13_6_12", Array(dteDate, DateAdd("m", 6, dteDate)), "Step 3"
'    dteDate = DateAdd("m", -18, dteDate2)
'    RunProc "q_r13_6_13", Array(dteDate, DateAdd("m", 6, dteDate)), "Step 4"
'    RunProc "q_r13_6_14", Array(dteDate), "Step 5" '   Prior to 18 months
'
''    ** Process transferred in transactions
'    dteDate = DateAdd("m", -3, dteDate2)
'    RunProc "q_r13_6_11T", Array(dteDate, dteDate2), "Step 6"
'    dteDate = DateAdd("m", -12, dteDate2)
'    RunProc "q_r13_6_12T", Array(dteDate, DateAdd("m", 6, dteDate)), "Step 7"
'    dteDate = DateAdd("m", -18, dteDate2)
'    RunProc "q_r13_6_13T", Array(dteDate, DateAdd("m", 6, dteDate)), "Step 8"
'    RunProc "q_r13_6_14T", Array(dteDate), "Step 8"    '   Prior to 18 months
'
'    Set rs = New ADODB.Recordset
'    GetDynamicRecordset "TEMP_r13_6_1", enTable, Array(), "Step 9", rs
'    Do While Not rs.EOF
'        With rs
'            iBal = 0
'            If !StockBal >= .Fields("6Month") Then
'                !Stck6Month = .Fields("6Month")
'            Else
'                !Stck6Month = iBal
'            End If
'            iBal = !StockBal - .Fields("6Month")
'            If iBal < 0 Then iBal = 0
'            If iBal >= .Fields("12Month") Then
'                !Stck12Month = .Fields("12Month")
'            Else
'                !Stck12Month = iBal
'            End If
'            iBal = iBal - .Fields("12Month")
'            If iBal < 0 Then iBal = 0
'            If iBal >= .Fields("18Month") Then
'                !Stck18Month = .Fields("18Month")
'            Else
'                !Stck18Month = iBal
'            End If
'            iBal = iBal - .Fields("18Month")
'            If iBal < 0 Then iBal = 0
'             If iBal >= .Fields("18MonthPlus") Then
'                !Stck18MonthPlus = .Fields("18MonthPlus")
'            Else
'                !Stck18MonthPlus = iBal
'            End If
'            iBal = iBal - .Fields("18MonthPlus")
'            If iBal < 0 Then iBal = 0
'
'            .Update
'            .MoveNext
'        End With
'    Loop
'    If rs.State = adStateOpen Then rs.Close
'    Set rs = Nothing
'
''Update product table
'    Set rs = New ADODB.Recordset
'    RunProc "q_StockAgeFigures", Array(dteDate2), "Step 10"
'
'EXIT_Handler:
'    If rs.State = adStateOpen Then rs.Close
'    Set rs = Nothing
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.UpdateStockAgeFigures"
'End Sub

Private Sub Class_Initialize()
    On Error GoTo errHandler
'    Set objOp = New a_Operation
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.Class_Initialize"
End Sub

Public Sub MarkOrdersAsCancelled(pSince As Date)
    On Error GoTo errHandler
Dim oUtil As z_SQL
Dim bAnswer As Boolean
    bAnswer = True
'    If pSince > DateAdd("m", -6, Date) Then
'        RaiseEvent OrderStatus("No", bAnswer)
'    Else
    If pSince > DateAdd("yyyy", -1, Date) Then
        RaiseEvent OrderStatus("Warning", bAnswer)
    Else
        RaiseEvent OrderStatus("", bAnswer)
    End If
    If bAnswer <> False Then
        RunProc "sp_MarkOldOrdersAsCancelled", Array(pSince), "", oPC.COShort
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.MarkOrdersAsCancelled(pSince)", pSince
End Sub
Public Sub MarkCustOrdersAsCancelled(pSince As Date)
    On Error GoTo errHandler
Dim oUtil As z_SQL
Dim bAnswer As Boolean
    bAnswer = True
'    If pSince > DateAdd("m", -6, Date) Then
'        RaiseEvent OrderStatus("No", bAnswer)
'    Else
    If pSince > DateAdd("yyyy", -1, Date) Then
        RaiseEvent OrderStatus("Warning", bAnswer)
    Else
        RaiseEvent OrderStatus("", bAnswer)
    End If
    If bAnswer <> False Then
        RunProc "sp_MarkOldCustOrdersAsCancelled", Array(pSince), ""
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.MarkCustOrdersAsCancelled(pSince)", pSince
End Sub
Public Sub MarkApprosAsCancelled(pSince As Date)
    On Error GoTo errHandler
Dim oUtil As z_SQL
Dim bAnswer As Boolean
    bAnswer = True
'    If pSince > DateAdd("m", -6, Date) Then
'        RaiseEvent OrderStatus("No", bAnswer)
'    Else
    If pSince > DateAdd("yyyy", -1, Date) Then
        RaiseEvent OrderStatus("Warning", bAnswer)
    Else
        RaiseEvent OrderStatus("", bAnswer)
    End If
    If bAnswer <> False Then
        RunProc "sp_MarkOldApprosAsCancelled", Array(pSince), ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.MarkApprosAsCancelled(pSince)", pSince
End Sub

Function DropTable(pName, pMsg As String, Optional pConn As ADODB.Connection) As Long
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim strSQL As String
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.COShort
    Else
        cmd.ActiveConnection = pConn
    End If
    strSQL = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[" & pName & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) drop table [dbo].[" & pName & "]"
    cmd.CommandText = strSQL
    cmd.CommandType = adCmdText
    RaiseEvent STATUS(strMMsg & ": " & pMsg)
    cmd.execute
    
EXIT_Handler:
    DropTable = 0
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.DropTable(pName,pMsg,pConn)", Array(pName, pMsg, pConn)
End Function
Function RunProc(pName As String, pParams As Variant, pMsg As String, Optional pConn As ADODB.Connection) As Long
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.COShort
    Else
        cmd.ActiveConnection = pConn
    End If
    cmd.CommandText = pName
    cmd.CommandType = adCmdStoredProc
    RaiseEvent STATUS(strMMsg & ": " & pMsg)
    If IsArray(pParams) Then
        If UBound(pParams, 1) > -1 Then
            retval = cmd.execute(, pParams)
        Else
            cmd.execute
        End If
    Else
        cmd.execute
    End If
    
    Set cmd = Nothing
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.RunProc(pName,pParams,pMsg,pConn)", Array(pName, pParams, pMsg, pConn)
End Function
Function RunSQL(pSQL As String, Optional pConn As ADODB.Connection) As Long
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim lngRecordsAffected As Long

Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.COShort
    Else
        cmd.ActiveConnection = pConn
    End If
    cmd.CommandText = pSQL
    cmd.CommandType = adCmdText
    
    cmd.execute lngRecordsAffected
    Set cmd = Nothing
    RunSQL = lngRecordsAffected
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.RunSQL(pSQL,pConn)", Array(pSQL, pConn)
End Function
Function RunGetRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset, Optional pConn As ADODB.Connection) As Long
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim lngRecordsAffected As Long
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set cmd = New ADODB.Command
    If pConn Is Nothing Then
        cmd.ActiveConnection = oPC.COShort
    Else
        cmd.ActiveConnection = pConn
    End If
    cmd.CommandText = pName
    cmd.CommandType = pType
    cmd.CommandTimeout = 8
    RaiseEvent STATUS(strMMsg & ": " & pMsg)
    If IsArray(pParams) Then
        If UBound(pParams, 1) > -1 Then
            Set rs = cmd.execute(, pParams)
        Else
            Set rs = cmd.execute()
        End If
    Else
        Set rs = cmd.execute()
    End If
    Set cmd = Nothing
    RunGetRecordset = rs.RecordCount
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.RunGetRecordset(pName,pType,pParams,pMsg,rs,pConn)", Array(pName, pType, pParams, _
         pMsg, rs, pConn)
End Function

Function GetDynamicRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset, _
                    Optional pConn As ADODB.Connection)
    On Error GoTo errHandler
    If pConn Is Nothing Then
        rs.Open pName, oPC.CO, adOpenDynamic, adLockOptimistic
    Else
        rs.Open pName, pConn, adOpenDynamic, adLockOptimistic
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.GetDynamicRecordset(pName,pType,pParams,pMsg,rs,pConn)", Array(pName, pType, _
         pParams, pMsg, rs, pConn)
End Function

Function UpdateFromBookfind(pAuthor As Boolean, _
pTitle As Boolean, _
pSubtitle As Boolean, _
pAvailability As Boolean, _
pBindingcode As Boolean, _
pEdition As Boolean, _
pSUpplierCode As Boolean, _
pPublisherName As Boolean, _
pSeriesTitle As Boolean, _
pPublicationDate As Boolean, _
pUKPrice As Boolean, _
pRRP As Boolean, _
pBIC As Boolean, _
pBookStatus As Boolean, _
Optional pSTAFFID As Long, _
Optional pEAN As String)
Dim OpenResult As Integer

10        On Error GoTo errHandler
      Dim strArticle As String
      Dim strTitleNet As String

      '     This is used when a new CD arrives.
      '     It finds all records on tblProduct where there are no details except ISBN and looks on the CD
      '     for details, if they are found then the record is updated.
      Dim dteStarted As Date
          Dim rs As ADODB.Recordset
      Dim strCode As String
          Dim oBF As a_BookFind
          Dim strMsg As String
          Dim StartTime
          Dim X As Long
          Dim iCancelled As Integer
          Dim strShortname As String
          Dim lngProgress As Long
          Dim lngMax As Long
20        iCancelled = False
          
          
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
          
30            RaiseEvent STATUS("Selecting records to update . . .")
40            DoEvents
50            Set rs = New ADODB.Recordset
60            rs.CursorLocation = adUseClient
70            If pEAN > "" Then
80                rs.Open "SELECT * FROM vBooksWithISBNs WHERE P_EAN > pEAN", oPC.COShort, adOpenKeyset, adLockOptimistic
90            Else
100               rs.Open "SELECT * FROM vBooksWithISBNs WHERE ISNULL(P_SKIPBFWASH,0) <> 1", oPC.COShort, adOpenKeyset, adLockOptimistic
110           End If
120           rs.MoveLast
130           lngProgress = 0
140           lngMax = rs.RecordCount
150           RaiseEvent Progress(lngProgress, lngMax)
160           dteStarted = Now()
170           RaiseEvent STATUS("Updating from CD . . ." & "Started:" & Format(dteStarted, "dd/mm/yyyy Hh:Nn") & ", Records: " & rs.RecordCount)
180           StartTime = Now
190           rs.MoveFirst
200           Set oBF = New a_BookFind
210       Do While Not rs.EOF
220           DoEvents
230           lngProgress = lngProgress + 1
240           RaiseEvent Progress(lngProgress, lngMax)
250           If Not (IsISBN10(FNS(rs![P_Code])) Or IsISBN13(FNS(rs![P_EAN]))) Then
260               GoTo ENDOFLOOP
270           End If
280           strCode = CStr(rs![P_EAN])
290           If strCode = "" Then strCode = CStr(rs![P_Code])
300           If oBF.FetchFromBF(strCode) = 0 Then
310               If pAuthor Then rs![P_MainAuthor] = Left$(FNS(oBF.MainAuthor), rs.Fields("P_MainAuthor").DefinedSize)
320               rs![P_Code] = Left$(FNS(oBF.Code), rs.Fields("P_Code").DefinedSize)
330               If pTitle Then
340                   StripArticle FNS(Left$(FNS(oBF.Title), rs.Fields("P_Title").DefinedSize)), strArticle, strTitleNet
350                   rs!P_Title = strTitleNet
360                   rs!P_Article = strArticle
370               End If
380               If pSubtitle Then rs![P_SubTitle] = Left$(FNS(oBF.SubTitle), rs.Fields("P_SubTitle").DefinedSize)
390               If pAvailability Then rs![P_Status] = Left(FNS(oBF.Availability), 1)   ' rs.Fields("P_SubTitle").DefinedSize)
400               If pBindingcode Then rs![P_Bindingcode] = Left$(FNS(oBF.BindingCode), rs.Fields("P_Bindingcode").DefinedSize)
410               If pEdition Then rs![P_Edition] = Left$(FNS(oBF.Edition), rs.Fields("P_Edition").DefinedSize)
420               If pPublisherName Then rs![P_Publisher] = Left$(FNS(oBF.PublisherName), rs.Fields("P_Publisher").DefinedSize)
430               If pSUpplierCode Then rs![P_BFSupplierCode] = Left$(FNS(oBF.DistributorCode), rs.Fields("P_BFSupplierCode").DefinedSize)
440                   strShortname = FNS(rs![P_Publisher])


      'We are fleshing out the tProduct table, and when there is a skeleton record (Defined as having no Publisher
      'name) then we fill in the other fields off Bookfind and then get the matching publisher record and see
      'whether it is on the TP table (we match by Publisher name). If not, we add a publisher record with
      'ISBNPrefix.
         'Get BF Publisher Name and ISBNPrefix
         'Try to match on papyrus TP_Shortname
         'if not matched then
         '   add new publisher record etc.
         'end if
         'assign new publisher record ID to Title
StartSearch2:
      '       If Not SeekRec("tblTradingPartner", "TP_Shortname", "=", strShortname) Then
      '            RetVal = GetMatchingBFPublisher(Trim$(BFRec.k1), ISBNPrefix, strImprintName, strPublisherName, strBFCode) 'returns ISBNPrefix in second param
      '            If RetVal = True Then  'found on BookFind''
      '
                                                  'This publisher is not on database , so add a record to
                                                  'tblTradingpartner and tblTradingPartner_Role and tblSupplier_Imprint
      '                 RetVal = AddAPublisher(strImprintName, strPublisherName, Trim$(ISBNPrefix), strBFCode, strShortname, 0)
      '                 GoTo StartSearch2
      '            End If
      '        Else
      '            rs![P_PublisherID] = gSeekRst![Trading_Partner_ID]
      '        End If
450                   If pUKPrice Then
460                       If IsNumeric(oBF.UKPrice) Then
470                           rs![P_UKPrice] = CCur(oBF.UKPrice) * oPC.Configuration.DefaultCurrency.Divisor
480                       Else
490                           rs![P_UKPrice] = Null
500                       End If
510                       If IsNumeric(oBF.USPrice) Then
520                           rs![P_USPrice] = CCur(oBF.USPrice) * oPC.Configuration.DefaultCurrency.Divisor
530                       Else
540                           rs![P_USPrice] = Null
550                       End If
560                   End If
570                   If pRRP Then
580                       If IsNumeric(oBF.LocalPrice) Then
590                           If oBF.LocalPrice > 0 Then
600                               rs![P_RRP] = FNN(oBF.LocalPrice * oPC.Configuration.DefaultCurrency.Divisor)
610                           Else
620                               rs![P_RRP] = FNN(oBF.UKPrice / IIf(oPC.Configuration.Currencies.FindBySysname("GBP") Is Nothing, 10, oPC.Configuration.Currencies.FindBySysname("GBP").Factor) * oPC.Configuration.DefaultCurrency.Divisor)
630                           End If
640                       End If
650                   End If
660                   If pSeriesTitle Then rs![P_SeriesTitle] = Left$(FNS(oBF.SeriesTitle), rs.Fields("P_SeriesTitle").DefinedSize)
670                   If pPublicationDate Then rs![P_PubDate] = Left$(FNS(oBF.PublicationDate), rs.Fields("P_Pubdate").DefinedSize)
680                   If pBIC Then rs![P_BIC] = Left$(FNS(oBF.BFClassification), rs![P_BIC].DefinedSize)
                   '  On Error Resume Next
690                   rs.Update
                    '  On Error GoTo errHandler
700               End If
710               If pBookStatus Then
720                   If DateDiff("yyyy", FND(rs!P_LastDateDelivered), Date) > 2 Then 'mark books received last more than 2 years ago and not on Bookfind as Out of Print
730                       rs!P_Status = "O"
740                       rs.Update
750                   End If
760               End If
                  
              'exit if operation cancelled
ENDOFLOOP:
            '  rs.CancelUpdate
770           rs.MoveNext
780           DoEvents
790       Loop

800       rs.Close
810       If pEAN = "" Then
820           oPC.COShort.execute "INSERT INTO tOPERATION (OP_STARTEDAT,OP_ENDEDAT,OP_NOMINALDATE,OP_TYPE,OP_STARTEDBYID,OP_RESULT) VALUES ('" & ReverseDateTime(dteStarted) & "','" & ReverseDateTime(Now) & "','" & ReverseDate(Date) & "'," & BookDataWash & "," & pSTAFFID & ",1)"
          
830           RaiseEvent STATUS("Updating from CD . . . complete.")
840       End If
850       Exit Function

ErrH::
860       MsgBox Error
870       Exit Function
880       Resume
890       Exit Function
errHandler:
900       ErrPreserve
910       If Err.Description = "Duplicate CODE" Or Err.Description = "Duplicate EAN" Then
920           rs.CancelUpdate
930           Resume ENDOFLOOP
940       End If
950       If ErrMustStop Then Debug.Assert False: Resume
960       ErrorIn "z_Batch.UpdateFromBookfind(pAuthor,pTitle,pSubtitle,pAvailability,pBindingcode,pEdition," & _
              "pSUpplierCode,pPublisherName,pSeriesTitle,pPublicationDate,pUKPrice,pRRP,pBIC,pBookStatus,pStaffID)", _
               Array(pAuthor, pTitle, pSubtitle, pAvailability, pBindingcode, pEdition, pSUpplierCode, _
               pPublisherName, pSeriesTitle, pPublicationDate, pUKPrice, pRRP, pBIC, pBookStatus, pSTAFFID)
End Function

Function createpublisherlist()
    On Error GoTo errHandler
Dim Nam
Dim Code
Dim extra As String
Dim strShortname
Dim iReccount As Long
Dim iRecsFound As Long
Dim oBF As a_BookFind
Dim rs As ADODB.Recordset
Dim i As Long
Dim strSearchTabs As String
Dim strRetrieveTabs As String
Dim strTemp As String
Dim strResults As String
Dim strCode As String
Dim iresultsFound As Variant
    oPC.CO.execute "DELETE FROM tDISTR"
    Set oBF = New a_BookFind
    oBF.InitializeScan "L2", "A", iresultsFound
    oPC.BF.DefineOutputFormat ("L2")
    i = 1
    Do While i <= 999999
        strCode = oBF.GetNextTerm
        If strCode > "" Then
            oPC.CO.execute "INSERT INTO tDISTR (DISTR_ID) VALUES ('" & strCode & "')"
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tDISTR", oPC.CO
    oPC.BF.DefineOutputFormat ("DN,AD,AP,TL,FX,EM")
    Do While Not rs.EOF
        strTemp = "FIND" & Chr(9) & "L2" & Chr(9) & FNS(rs!DISTR_ID)
        Check oBF.Search(strTemp, iRecsFound) = 0, EXC_GENERAL, "Error in oBF:Search"
        If iRecsFound > 0 Then
            Check oBF.GetRecordByNum(1) = 0, EXC_GENERAL, "Error in oBF:GetRecordByNum"
            oBF.LoadDistrRec (1)
            oPC.CO.execute "UPDATE tDISTR SET DISTR_NAME = '" & SQLQuotes(oBF.DistributorName) & "' WHERE DISTR_ID = '" & FNS(rs!DISTR_ID) & "'"
        End If
        rs.MoveNext
    Loop
    rs.Close
    Set oBF = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.createpublisherlist"
End Function
Function bfPubSearch(fld As String, arg As String) As Long
    On Error GoTo errHandler
Dim iErr As Long
Dim strSearchSpec As String
Dim iRecsFound As Long
    ChDrive "C:"
    ChDir "C:\Bookfind"

    strSearchSpec = "FIND" & Chr$(9) & fld & Chr$(9) & arg
    Check oPC.BF.Search(strSearchSpec, iRecsFound), EXC_GENERAL, "Error in BF.Search"
    bfPubSearch = iRecsFound
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.bfPubSearch(fld,arg)", Array(fld, arg)
End Function

Public Function CustomerIndexClashes(lngTPID As Long, strACno As String) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CustomerClash"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TPID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = lngTPID
    Set par = cmd.CreateParameter("@ACNO", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = strACno
    Set par = cmd.CreateParameter("@R", adInteger, adParamOutput)
    cmd.Parameters.Append par
    par.Value = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CustomerIndexClashes = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function


    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.CustomerIndexClashes(lngTPID,strACno)", Array(lngTPID, strACno)
End Function

Public Function ExportDebtorsTrading_Pastel(SMID As Long, pLASTTRIDExported As Long, dteDateSince As Date, lngPeriod As Long) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String

Dim Res As Boolean

Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDB
'-------------------------------

'Extract data
    Set cmd = New ADODB.Command
    cmd.CommandText = "ExportDebtorsTrading_Pastel"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@SMID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = SMID
    
    Set par = cmd.CreateParameter("@LASTTRIDExported", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pLASTTRIDExported
    
    Set par = cmd.CreateParameter("@DATESINCE", adDate, adParamInput)
    cmd.Parameters.Append par
    par.Value = ReverseDate(dteDateSince)
    
    Set par = cmd.CreateParameter("@PERIOD", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = lngPeriod
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamInputOutput)
    cmd.Parameters.Append par
    par.Value = 0
    
    cmd.ActiveConnection = oPC.CO
    cmd.CommandTimeout = 0
    cmd.execute
    ExportDebtorsTrading_Pastel = (cmd.Parameters(4) = 0)
    Set cmd = Nothing
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ExportDebtorsTrading_Pastel(SMID,pLASTTRIDExported,dteDateSince)", Array(SMID, _
         pLASTTRIDExported, dteDateSince)
End Function
Public Function ExportDebtorsTrading_Pastel_Confirmed()
Dim strSQL As String
Dim strCommand As String
Dim Res As Boolean
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol

'Prepare folders
    dte = Now()
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\ACCOUNTING") Then
        fol = fs.CreateFolder(oPC.SharedFolderRoot & "\ACCOUNTING")
    End If

    strFilePath = oPC.SharedFolderRoot & "\ACCOUNTING\SALES_INV_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    'Takes all from table tPASTEL
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vPASTELExport_TAXINVOICES"
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -eBCPError.sal -c -t\, -q  -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 20000)

    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vPASTELExport_TAXCREDITNOTES"
    strCommand = "bcp """ & strSQL & """ queryout """ & Replace(strFilePath, "_INV_", "_CN_") & """ -eBCPError.sal -c -t\, -q  -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand, vbHide, 20000)

End Function
Public Function ExportCreditorsTrading_Pastel_Confirmed()
    On Error GoTo errHandler
Dim strSQL As String
Dim strCommand As String
Dim Res As Boolean
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol


    dte = Now()
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\ACCOUNTING") Then
        fol = fs.CreateFolder(oPC.SharedFolderRoot & "\ACCOUNTING")
    End If

    strFilePath = oPC.SharedFolderRoot & "\ACCOUNTING\CREDITORS_PURCHASES_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vPASTELExport_Purchases"
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -eAccountingExportPurchases.txt -c -t\, -q  -Usa -Pcar -S" & oPC.servername
   F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 25000
'Shell strCommand
    strFilePath = oPC.SharedFolderRoot & "\ACCOUNTING\CREDITORS_RETURN_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vPASTELExport_PurchaseReturns"
    strCommand = "bcp """ & strSQL & """ queryout """ & Replace(strFilePath, "_DB_", "_CR_") & """ -eAccountingExportReturns.txt -c -t\, -q  -Usa -Pcar -S" & oPC.servername
    F_7_AB_1_ShellAndWaitSimple strCommand, vbHide, 25000
'Shell strCommand

    Exit Function

    Exit Function
errHandler:
    ErrorIn "z_Batch.ExportCreditorsTrading_Pastel_Confirmed"
End Function

Public Function ExportCreditorsTrading_Pastel(SMID As Long, pLASTTRIDExported As Long, dteDateSince As Date, lngPeriod As Long) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol
Dim Res As Boolean

Dim OpenResult As Integer

'-------------------------------
    OpenResult = oPC.OpenDB
'-------------------------------
'Prepare folders
    dte = Now()
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\ACCOUNTING") Then
        fol = fs.CreateFolder(oPC.SharedFolderRoot & "\ACCOUNTING")
    End If


'Extract data
    Set cmd = New ADODB.Command
    cmd.CommandText = "ExportCreditorsTrading_Pastel"
    cmd.CommandType = adCmdStoredProc
    
    
    Set par = cmd.CreateParameter("@SMID", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = SMID
    
    Set par = cmd.CreateParameter("@LASTTRIDExported", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = pLASTTRIDExported
    
    Set par = cmd.CreateParameter("@DATESINCE", adDate, adParamInput)
    cmd.Parameters.Append par
    par.Value = ReverseDate(dteDateSince)
    
    Set par = cmd.CreateParameter("@PERIOD", adInteger, adParamInput)
    cmd.Parameters.Append par
    par.Value = lngPeriod
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamInputOutput)
    cmd.Parameters.Append par
    par.Value = 0
    
    cmd.ActiveConnection = oPC.CO
    cmd.CommandTimeout = 0
    cmd.execute
    
    ExportCreditorsTrading_Pastel = (cmd.Parameters(4) = 0)
    Set cmd = Nothing
    

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ExportCreditorsTrading_Pastel(SMID,pLASTTRIDExported,dteDateSince)", Array(SMID, _
         pLASTTRIDExported, dteDateSince)
End Function
Public Function SaveExportToPastel(DRCR As String) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol
Dim Res As Boolean

    Set cmd = New ADODB.Command
    cmd.CommandText = "SaveExportToPastel"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Type", adVarChar, adParamInput, 2, DRCR)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.CO
    cmd.CommandTimeout = 0
    cmd.execute
    
    'ExportCreditorsTrading_PastelToSTore = (cmd.Parameters(4) = 0)
    Set cmd = Nothing
'
    Exit Function
errHandler:
    ErrorIn "z_Batch.SaveExportToPastel"
End Function

Public Function ExportCustomers_Pastel(SMID As Long, dteDateSince As Date) As Boolean
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol


'Extract data
    dte = Now()
    Set cmd = New ADODB.Command
    cmd.CommandText = "ExportCustomers_Pastel"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@SMID", adInteger, adParamInput, , SMID)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@DATESINCE", adDate, adParamInput)
    par.Value = ReverseDate(dteDateSince)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamInputOutput)
    cmd.Parameters.Append par
    par.Value = 0
    
    cmd.ActiveConnection = oPC.CO
    cmd.execute
    
    ExportCustomers_Pastel = (cmd.Parameters(2) = 0)
    Set cmd = Nothing

'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.ExportCustomers_Pastel(SMID,dteDateSince)", Array(SMID, dteDateSince)
'errHandler:
'    ErrorIn "z_Batch.ExportCreditorsTrading_PastelToSTore"
    Exit Function
errHandler:
    ErrorIn "z_Batch.ExportCustomers_Pastel"
End Function
Public Function ExportCustomers_Pastel2() As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim strErrorPath As String
Dim dte As Date
Dim fol
Dim Res As Boolean

    dte = Now()
    strFilePath = oPC.SharedFolderRoot & "\ACCOUNTING\CUST_EXPORT_" & Format(dte, "yyyymmddHHNN") & ".txt"
    strErrorPath = Replace(strFilePath, ".txt", ".ERR")
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    If fs.FileExists(strErrorPath) Then
        fs.DeleteFile strErrorPath, True
    End If

    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vPASTELExport_CUST"
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -e """ & strErrorPath & """ -c -t\, -q  -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ExportCustomers_Pastel2"
End Function
Public Function ExportSuppliers_Pastel(SMID As Long, dteDateSince As Date) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol

'Prepare folders
    dte = Now()
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\ACCOUNTING") Then
        fol = fs.CreateFolder(oPC.SharedFolderRoot & "\ACCOUNTING")
    End If

    strFilePath = oPC.SharedFolderRoot & "\ACCOUNTING\SUPP_EXPORT" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If

'Extract data
    Set cmd = New ADODB.Command
    cmd.CommandText = "ExportSuppliers_Pastel"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@SMID", adInteger, adParamInput, , SMID)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@DATESINCE", adDate, adParamInput)
    'If dteDateSince > 0 Then
        par.Value = ReverseDate(dteDateSince)
    'End If
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@RES", adInteger, adParamInputOutput)
    cmd.Parameters.Append par
    par.Value = 0
    
    cmd.ActiveConnection = oPC.CO
    cmd.execute
    
    ExportSuppliers_Pastel = (cmd.Parameters(2) = 0)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ExportSuppliers_Pastel(SMID,dteDateSince)", Array(SMID, dteDateSince)
End Function
Public Function ExportSuppliers_Pastel2() As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim strErrorPath As String
Dim dte As Date
Dim fol
Dim Res As Boolean

    dte = Now()
    strFilePath = oPC.SharedFolderRoot & "\ACCOUNTING\SUPP_EXPORT_" & Format(dte, "yyyymmddHHNN") & ".txt"
    strErrorPath = Replace(strFilePath, ".txt", ".ERR")
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    If fs.FileExists(strErrorPath) Then
        fs.DeleteFile strErrorPath, True
    End If

    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vPASTELExport_SUPP"
  
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -e """ & strErrorPath & """ -c -t\, -q  -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ExportCustomers_Pastel2"
End Function
Public Function ImportDebtorsME(SMID As Long, pFilePath As String)
    On Error GoTo errHandler
Dim strCommand As String
Dim oF As New z_TextFile
Dim strIn As String
Dim ar() As String
Dim rs As ADODB.Recordset

    oPC.CO.execute "Delete FROM tDEBTORS_IMPORT"
    oF.OpenTextFileToRead pFilePath
    strIn = oF.ReadLinefromTextFile
    Set rs = New ADODB.Recordset
    rs.Open "Select * from tDEBTORS_IMPORT", oPC.CO, adOpenDynamic, adLockOptimistic
    
    Do While Not strIn = ""
        ar = Split(strIn, ",")
        rs.AddNew
            rs.Fields("DI_ACNO") = FNS(ar(0))
            rs.Fields("DI_TOTALBAL") = FNDBL(ar(1))
            rs.Fields("DI_30") = FNDBL(ar(2))
            rs.Fields("DI_60") = FNDBL(ar(3))
            rs.Fields("DI_90") = FNDBL(ar(4))
            rs.Fields("DI_120") = FNDBL(ar(5))
            rs.Fields("DI_INTEREST") = FNDBL(ar(6))
            rs.Fields("DI_BLOCKED") = FNB(ar(7))
            rs.Fields("DI_Action") = "1"
        rs.Update
        If oF.IsEOF Then
            strIn = ""
        Else
            strIn = oF.ReadLinefromTextFile
        End If
    Loop
    rs.Close
    Set rs = Nothing
    oF.CloseTextFile
    RunSQL "UPDATE tDEBTORS_IMPORT SET DI_TPID = TP_ID FROM tDEBTORS_IMPORT JOIN tTP ON DI_ACNO = TP_ACNO", oPC.CO
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ImportDebtorsME(SMID,pFilePath)", Array(SMID, pFilePath)
End Function
Public Function ImportWebMasterList(pFilePath As String)
    On Error GoTo errHandler
Dim strCommand As String
Dim oF As New z_TextFile
Dim strIn As String
Dim ar() As String
Dim rs As ADODB.Recordset

    oPC.CO.execute "Truncate Table tWebMasterList"
    oF.OpenTextFileToRead pFilePath
    strIn = oF.ReadLinefromTextFile
    Set rs = New ADODB.Recordset
    rs.Open "Select * from tWebMasterList", oPC.CO, adOpenDynamic, adLockOptimistic
    
    Do While Not strIn = ""
        ar = Split(strIn, ",")
        rs.AddNew
            rs.Fields("WM_EAN_REQ") = FNS(ar(0))
        rs.Update
        If oF.IsEOF Then
            strIn = ""
        Else
            strIn = oF.ReadLinefromTextFile
        End If
    Loop
    rs.Close
    Set rs = Nothing
    oF.CloseTextFile
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ImportWebMasterList(pFilePath)", Array(pFilePath)
End Function

Public Function CreateBFBALTransaction(pDate As Date, pPeriodDescription As String) As Integer
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "CreateBFBALTransaction"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.CO
    
    Set par = cmd.CreateParameter("@DTE", adDate, adParamInput)
    cmd.Parameters.Append par
    par.Value = pDate
    
    Set par = cmd.CreateParameter("@PERIODDESCRIPTION", adVarChar, adParamInput, 20)
    cmd.Parameters.Append par
    par.Value = pPeriodDescription
    
    cmd.execute
    
    CreateBFBALTransaction = 0
    Set cmd = Nothing

    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.CreateBFBALTransaction(pDate,pPeriodDescription)", Array(pDate, _
         pPeriodDescription)
End Function

Public Function ImportCustomers_Pastel(SMID As Long, pFilePath As String)
    On Error GoTo errHandler
Dim strCommand As String
Dim oF As New z_TextFile
Dim strIn As String
Dim ar() As String
Dim rs As ADODB.Recordset
Dim rsS As ADODB.Recordset

    oPC.CO.execute "Delete FROM tPASTEL_CUST_IMPORT"
    oF.OpenTextFileToRead pFilePath
    strIn = oF.ReadLinefromTextFile
    Set rs = New ADODB.Recordset
    rs.Open "Select * from tPASTEL_CUST_IMPORT", oPC.CO, adOpenDynamic, adLockOptimistic
    Do While Not strIn = ""
        ar = Split(strIn, """,""")
            rs.AddNew
            rs.Fields("PC_ACNO") = StripEnclosingQuotes(ar(0))
            rs.Fields("PC_NAME") = StripEnclosingQuotes(ar(1))
            rs.Fields("PC_ADD1") = StripEnclosingQuotes(ar(2))
            rs.Fields("PC_ADD2") = StripEnclosingQuotes(ar(3))
            rs.Fields("PC_ADD3") = StripEnclosingQuotes(ar(4))
            rs.Fields("PC_ADD4") = StripEnclosingQuotes(ar(5))
            rs.Fields("PC_ADD5") = StripEnclosingQuotes(ar(6))
            rs.Fields("PC_TEL") = StripEnclosingQuotes(ar(7))
            rs.Fields("PC_FAX") = StripEnclosingQuotes(ar(8))
            rs.Fields("PC_CONTACTPERSON") = StripEnclosingQuotes(ar(9))
            rs.Fields("PC_DUMMY1") = StripEnclosingQuotes(ar(10))
            rs.Fields("PC_DUMMY2") = StripEnclosingQuotes(ar(11))
            rs.Fields("PC_DUMMY3") = StripEnclosingQuotes(ar(12))
            rs.Fields("PC_DUMMY4") = StripEnclosingQuotes(ar(13))
            rs.Fields("PC_DADD1") = StripEnclosingQuotes(ar(14))
            rs.Fields("PC_DADD2") = StripEnclosingQuotes(ar(15))
            rs.Fields("PC_DADD3") = StripEnclosingQuotes(ar(16))
            rs.Fields("PC_DADD4") = StripEnclosingQuotes(ar(17))
            rs.Fields("PC_DADD5") = StripEnclosingQuotes(ar(18))
            rs.Fields("PC_BLOCKED") = IIf(StripEnclosingQuotes(ar(19)) = "Y", 1, 0)
            rs.Fields("PC_DISCOUNT") = StripEnclosingQuotes(ar(20))
            rs.Fields("PC_DUMMY5") = StripEnclosingQuotes(ar(21))
            rs.Fields("PC_DUMMY6") = StripEnclosingQuotes(ar(22))
            rs.Fields("PC_DUMMY7") = StripEnclosingQuotes(ar(23))
            rs.Fields("PC_CATEGORY") = StripEnclosingQuotes(ar(24))
            rs.Fields("PC_CURRENCYCODE") = StripEnclosingQuotes(ar(25))
            rs.Fields("PC_TERMS") = StripEnclosingQuotes(ar(26))
            rs.Fields("PC_CREDITLIMIT") = StripEnclosingQuotes(ar(27))
            rs.Fields("PC_DUMMY8") = StripEnclosingQuotes(ar(28))
            rs.Fields("PC_DUMMY9") = StripEnclosingQuotes(ar(29))
            rs.Fields("PC_DUMMY10") = StripEnclosingQuotes(ar(30))
            rs.Fields("PC_DUMMY11") = StripEnclosingQuotes(ar(31))
            rs.Fields("PC_DUMMY12") = StripEnclosingQuotes(ar(32))
            rs.Fields("PC_TAXREFERENCE") = StripEnclosingQuotes(ar(33))
            rs.Fields("PC_DUMMY13") = StripEnclosingQuotes(ar(34))
            rs.Fields("PC_MOBILE") = StripEnclosingQuotes(ar(35))
            rs.Fields("PC_EMAIL") = StripEnclosingQuotes(ar(36))
            rs.Fields("PC_DUMMY14") = StripEnclosingQuotes(ar(37))
            rs.Fields("PC_DUMMY15") = StripEnclosingQuotes(ar(38))
            rs.Fields("PC_COUNTRYCODE") = StripEnclosingQuotes(ar(39))
            rs.Fields("PC_FREIGHT") = StripEnclosingQuotes(ar(40))
            rs.Fields("PC_DUMMY16") = StripEnclosingQuotes(ar(41))
            rs.Fields("PC_DUMMY17") = StripEnclosingQuotes(ar(42))
            rs.Fields("PC_DUMMY19") = StripEnclosingQuotes(ar(43))
            rs.Fields("PC_ACTION") = 1
        rs.Update
        If oF.IsEOF Then
            strIn = ""
        Else
            strIn = oF.ReadLinefromTextFile
        End If
    Loop
    rs.Close
    Set rs = Nothing
    oF.CloseTextFile
    
'
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ImportCustomers_Pastel(SMID,pFilePath)", Array(SMID, pFilePath)
End Function
Public Function ImportCustomers_Pastel2()
    On Error GoTo errHandler
Dim cmd As ADODB.Command
    
    Set cmd = New ADODB.Command
    cmd.CommandText = "Import_PASTEL_Customers"
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = oPC.CO
    cmd.execute
    Set cmd = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ImportCustomers_Pastel2"
End Function




Public Function ExportInventorySinceDate(dteDateSince As Date) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol
Dim fout, f
Dim strBuff As String
Dim strBuff2 As String
Dim strBuff3 As String
Dim Res As Boolean

'Prepare folders
    dte = Now()
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\FilesForExport") Then
        fol = fs.CreateFolder(oPC.SharedFolderRoot & "\FilesForExport")
    End If

    strFilePath = oPC.SharedFolderRoot & "\FilesForExport\INV_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vInventoryExportForWeb WHERE P_DATERECORDADDED >= '" & ReverseDate(dteDateSince) & "' or P_DATELASTMODIFIED >= '" & ReverseDate(dteDateSince) & "'"
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -eBCPError.sal -c -q -t, -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    strFilePath = oPC.SharedFolderRoot & "\FilesForExport\INV_Categories_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vSectionExportForWeb"
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -eBCPError.sal -c -q -t, -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ExportInventorySinceDate(dteDateSince)", dteDateSince
End Function

Public Function ExportInventoryByCategory(pDictID As Long) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol
Dim fout, f
Dim strBuff As String
Dim strBuff2 As String
Dim strBuff3 As String
Dim Res As Boolean

'Prepare folders
    dte = Now()
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\FilesForExport") Then
        fol = fs.CreateFolder(oPC.SharedFolderRoot & "\FilesForExport")
    End If

    strFilePath = oPC.SharedFolderRoot & "\FilesForExport\INVENTORY_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vInventoryExport WHERE DICT_ID = " & CStr(pDictID)
    
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -eBCPError.sal -c -q -t, -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    strFilePath = oPC.SharedFolderRoot & "\FilesForExport\INVENTORY_Categories_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vSectionExportForWeb"
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -eBCPError.sal -c -q -t, -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ExportInventoryByCategory(pDictID)", pDictID
End Function


Public Function ExportInventoryForWeb(pDictID As Long) As Boolean
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strSQL As String
Dim strCommand As String
Dim fs As New FileSystemObject
Dim strFilePath As String
Dim dte As Date
Dim fol
Dim fout, f
Dim strBuff As String
Dim strBuff2 As String
Dim strBuff3 As String
Dim Res As Boolean

'Prepare folders
    dte = Now()
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\FilesForExport") Then
        fol = fs.CreateFolder(oPC.SharedFolderRoot & "\FilesForExport")
    End If

    strFilePath = oPC.SharedFolderRoot & "\FilesForExport\INVENTORY_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    If pDictID <> 0 Then
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vInventoryExportForWEB WHERE DICT_ID = " & CStr(pDictID)
    Else
        strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vInventoryExportForWEB"
    End If
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -eBCPError.sal -c -q -t, -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    strFilePath = oPC.SharedFolderRoot & "\FilesForExport\INVENTORY_Categories_" & Format(dte, "yyyymmddHHNN") & ".txt"
    If fs.FileExists(strFilePath) Then
        fs.DeleteFile strFilePath, True
    End If
    strSQL = "SELECT * FROM " & oPC.DatabaseName & ".dbo.vSectionExportForWeb"
    strCommand = "bcp """ & strSQL & """ queryout """ & strFilePath & """ -eBCPError.sal -c -q -t, -Usa -Pcar -S " & oPC.servername
    Res = F_7_AB_1_ShellAndWaitSimple(strCommand)

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.ExportInventoryByCategory(pDictID)", pDictID
End Function

Public Sub DeleteObsoleteCustOrders(dteSince As Date, bDeleteOrphanedCustomers, lngOrdersDeleted As Long, lngCustomersDeleted As Long)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    If dteSince > DateAdd("yyyy", -1, Date) Then
        If MsgBox("The date you have specified is less than a year ago. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
        If MsgBox("Orders " & IIf(bDeleteOrphanedCustomers, "and customers ", "") & "prior to " & Format(dteSince, "dd/mm/yyyy") & " will be deleted. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
        
    Set cmd = New ADODB.Command
    cmd.CommandText = "DeleteObsoleteOrders"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Since", adDate, adParamInput, , dteSince)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@DATESINCE", adBoolean, , , bDeleteOrphanedCustomers)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@QtyOrdersDeleted", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@QtyCustomersDeleted", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.CO
    cmd.execute
    
    lngOrdersDeleted = CLng(cmd.Parameters(2))
    lngCustomersDeleted = CLng(cmd.Parameters(3))
    Set cmd = Nothing

    Exit Sub

End Sub

Public Sub DeleteObsoleteQuotations(dteSince As Date, bDeleteOrphanedCustomers, lngDocsDeleted As Long, lngCustomersDeleted As Long)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    If dteSince > DateAdd("yyyy", -1, Date) Then
        If MsgBox("The date you have specified is less than a year ago. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
        If MsgBox("Quotations " & IIf(bDeleteOrphanedCustomers, "and customers ", "") & "prior to " & Format(dteSince, "dd/mm/yyyy") & " will be deleted. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
        
    Set cmd = New ADODB.Command
    cmd.CommandText = "DeleteObsoleteQuotations"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Since", adDate, adParamInput, , dteSince)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@DATESINCE", adBoolean, , , bDeleteOrphanedCustomers)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@QtyOrdersDeleted", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@QtyCustomersDeleted", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.CO
    cmd.execute
    
    lngDocsDeleted = CLng(cmd.Parameters(2))
    lngCustomersDeleted = CLng(cmd.Parameters(3))
    Set cmd = Nothing

    Exit Sub

End Sub

Public Sub TransferOldinvoicesToCasual(dteSince As Date, bDeleteOrphanedCustomers, lngDocsDeleted As Long, lngCustomersDeleted As Long)
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
    If dteSince > DateAdd("yyyy", -1, Date) Then
        If MsgBox("The date you have specified is less than a year ago. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("Invoices dated prior to " & Format(dteSince, "dd/mm/yyyy") & " for casual customers will be reallocated to a/c 'Casual' and " & IIf(bDeleteOrphanedCustomers, "and orphaned customers will be deleted ", "") & " Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
        
    Set cmd = New ADODB.Command
    cmd.CommandText = "TransferOldinvoicesToCasual"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@Since", adDate, adParamInput, , dteSince)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@DATESINCE", adBoolean, , , bDeleteOrphanedCustomers)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@QtyOrdersDeleted", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    Set par = cmd.CreateParameter("@QtyCustomersDeleted", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.CO
    cmd.execute
    
    lngDocsDeleted = FNN(cmd.Parameters(2))
    lngCustomersDeleted = FNN(cmd.Parameters(3))
    Set cmd = Nothing

    Exit Sub

End Sub

