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
Attribute VB_Creatable = False
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
Public Event Status(Msg As String)
Public Event OrderStatus(RequestStatus As String, Answer As Boolean)
Public Event DEResult(pMsg As String, pErr As Long)
Dim WithEvents oCD As Z_CDWrite
Attribute oCD.VB_VarHelpID = -1
Dim mNominalDate As Date

Private Sub oCD_Status(pMsg As String, bError As Boolean)
    On Error GoTo errHandler
    RaiseEvent Status(pMsg)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.oCD_Status(pMsg,bError)", Array(pMsg, bError)
End Sub
Private Sub oCD_Complete()
    On Error GoTo errHandler
    DailyUpdate mNominalDate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.oCD_Complete"
End Sub

'Public Sub SetMainMsg(pMsg As String)
'    On Error GoTo errHandler
'    strMMsg = pMsg
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.SetMainMsg(pMsg)", pMsg
'End Sub
'Function TableExists(pTableName As String)
'    On Error GoTo errHandler
'Dim cmdADO As ADODB.Command
'
'    cmdADO.ActiveConnection =
'
'    TableExists = True
'EXIT_Handler:
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.TableExists(pTableName)", pTableName
'End Function
'Private Sub oCD_EraseDiscComplete()
'Dim strErrmsg As String
'    On Error GoTo errHandler
'    If oPC.CompressBackup = "YES" Then
'        oCD.AddFile oPC.BackupFolder & "PBKS.ZIP"
'        If oCD.WriteDisc(strErrmsg) = False Then
'            RaiseEvent DEResult("Error writing to CD.", 1)
'        End If
'    Else
'        oCD.AddFile oPC.BackupFolder & "PBKS.BAK"
'        oCD.AddFile oPC.BackupFolder & "PBKSMASTER.BAK"
'        If oCD.WriteDisc(strErrmsg) = False Then
'            RaiseEvent DEResult("Error writing to CD.", 1)
'        End If
'    End If
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.oCD_EraseDiscComplete"
'End Sub
'Public Function DailyUpdateincBackup(pNominalDate As Date)
'    On Error GoTo errHandler
'Dim fs As New FileSystemObject
'Dim strBUFolder As String
'Dim oBU As New z_PBKSBackup
'Dim strErrmsg As String
'
'    pBackupOK = True
'    mNominalDate = pNominalDate
'    RaiseEvent status("Backing up the database . . .")
'    If oBU.BackupToLocal = True Then
'        If Not (fs.FileExists(oPC.BackupFolder & "PBKS.BAK") And fs.FileExists(oPC.BackupFolder & "PBKSMASTER.BAK")) Then
'            pBackupOK = False
'            Exit Function
'        End If
'    End If
' '''''
'    If pBackupOK Then
'        If oPC.RemoveableBackupMedium = "CD" Then
'            If oPC.CompressBackup = "YES" Then
'                If oBU.ZIPBackupToLocal = False Then
'                    RaiseEvent DEResult("Zip operation to local drive failed.", 1)
'                    pBackupOK = False
'                    Exit Function
'                Else
'                    Set oCD = New Z_CDWrite
'                    If oCD.SetDefaultDrive("") = False Then
'                        RaiseEvent DEResult("No CD drive to write to.", 1)
'                        pBackupOK = False
'                        Exit Function
'                    Else
'                        oCD.EraseCD
'                    End If
'                End If
'            Else
'                oCD.EraseCD
'            End If
'        Else
'            If oPC.CompressBackup = "YES" Then
'                If oBU.ZIPBackupToNonLocal = False Then
'                    RaiseEvent DEResult("Zip operation to removeable drive failed.", 1)
'                    pBackupOK = False
'                    Exit Function  'No dayend if backup does not work
'                End If
'            Else
'                fs.CopyFile oPC.BackupFolder & "PBKS.BAK", oPC.RemovableBackupShareName, True
'                fs.CopyFile oPC.BackupFolder & "PBKSMASTER.BAK", oPC.RemovableBackupShareName, True
'            End If
'        End If
'        DailyUpdate (pNominalDate)
'    End If
' ''''''
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.DailyUpdateincBackup(pNominalDate,pBackupOK)", Array(pNominalDate, pBackupOK)
'End Function
'
Public Function DailyUpdate(pNominalDate As Date) As Long
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
Dim X As String, y As String
Dim f As File
Dim strSQL As String
Dim dteLastCopied As Date
Dim strParentFolderName As String
Dim retval
Dim rs As ADODB.Recordset
Dim prm As ADODB.Parameter
Dim strPosition As String

    RaiseEvent Status("Dayend running . . .")
    DoEvents
        
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    Set cmd.ActiveConnection = cnPapy
    cmd.CommandText = "sp_DAYEND"
    cmd.CommandType = adCmdStoredProc
    
'0
    Set prm = Nothing
    Set prm = New ADODB.Parameter
    prm.Type = adDate
    prm.Direction = adParamInput
    prm.Value = pNominalDate
    cmd.Parameters.Append prm
        
'3
    Set prm = Nothing
    Set prm = New ADODB.Parameter
    prm.Type = adInteger
    prm.Direction = adParamOutput
    prm.Value = lngResult
    cmd.Parameters.Append prm
    
'4
    Set prm = Nothing
    Set prm = New ADODB.Parameter
    prm.Type = adInteger
    prm.Direction = adParamOutput
    prm.Value = lngPosition  'Possible position of error if error occurs
    cmd.Parameters.Append prm
'5
    Set prm = Nothing
    Set prm = New ADODB.Parameter
    prm.Type = adInteger
    prm.Direction = adParamOutput
    prm.Value = lngOPID  'operation ID
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    
    cmd.Execute
    If cmd.Parameters(3) <> 0 Or IsNull(cmd.Parameters(1)) Then
        strPosition = CStr(cmd.Parameters(2))
        GoTo SP_ERR_Handler
    End If
'
    Set ps = New z_papysystem
    If Not IsNull(cmd.Parameters(5)) Then
        cnPapy.Execute "UPDATE tOPERATION SET OP_FULLREPORT = " & ps.GetDiskStatus() & " WHERE OP_ID = " & CLng(cmd.Parameters(3))
    End If
    Set ps = Nothing
    
'
'
'
'
''Run apend query to update the values in the StockSummary table.
'    strMMsg = "Stock summary"
'    LoadStockSummary
'
''Update stock age figures in product table
'    strMMsg = "Age stock"
'    UpdateStockAgeFigures


 '  oPC.CO.Close
EXIT_Handler:
    RaiseEvent DEResult("Dayend complete", 0)
    DoEvents
    Set cmd = Nothing
    Set ps = Nothing
    Exit Function

SP_ERR_Handler:
    RaiseEvent DEResult("The dayend has failed at position " & strPosition & ". Contact Wizards Software", 99)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.DailyUpdate(pNominalDate)", pNominalDate, , cnPapy
End Function
Public Function SetOOS()
    On Error GoTo errHandler

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Batch.SetOOS"
End Function
'Public Function RemoveSOs(pDate As Date, pContactID As Long)
'    On Error GoTo errHandler
'
'Dim oError As a_Error
'Dim strConn As String
'Dim objOp As a_Operation
'Dim cmd As ADODB.Command
'
'    Set cmd = New ADODB.Command
'    Set objOp = New a_Operation
'    objOp.BeginEdit
'        fMode = "ADD"
'        objOp.StartedAt = Now()
'        objOp.TypeID = 3
'        objOp.NominalDate = pDate
'        objOp.OperatorID = pContactID
'    objOp.ApplyEdit lngResult
'    cmd.CommandText = "RemoveOLDSOs_1"
'    cmd.CommandType = adCmdStoredProc
'    cmd.Execute , Array(pDate)
'    retval = RunProc("RemoveOLDSOs_2", "", "", oPC.CO)
'   ' retval = RunProc("RemoveOLDSOs_3", oPC.CO, "") 'cannot run this as it deletes the links between product and matching tables
'EXIT_Handler:
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.RemoveSOs(pDate,pContactID)", Array(pDate, pContactID)
'End Function
'
'
'Function RebuildPublisherTable()
'    On Error GoTo errHandler
'Dim strSQL As String
'
'    strSQL = "SELECT  P_Publisher INTO PublisherLIst_TEMP FROM tPRODUCT GROUP BY p_Publisher"
'    retval = Me.DropTable("PublisherLIst_TEMP", "Erasing Table  . . . ")
'
'    Me.RunSQL strSQL, oPC.CO '"q_ListOfPublishers", Array(), "Building list of Publishers  . . . "
'
'    SaveSetting "SupplierChange", "Settings", "LastPubListRebuild", Now
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.RebuildPublisherTable"
'End Function
'
'Function SupplierChange(pcolPublisher As Collection, pLastTPID, pNewTPID, pMsg As String, Optional pConn As ADODB.Connection) As Long
'    On Error GoTo errHandler
'
'
'Dim cmd As ADODB.Command
'Dim lngRecordsAffected As Long
'Dim strSQL As String
'Dim strWHERE As String
'Dim i As Integer
'
'    Set cmd = New ADODB.Command
'    If pConn Is Nothing Then
'        cmd.ActiveConnection = oPC.CO
'    Else
'        cmd.ActiveConnection = pConn
'    End If
'
'    strWHERE = ""
'    For i = 1 To pcolPublisher.Count
'        If Len(strWHERE) > 0 Then
'            strWHERE = strWHERE & " OR "
'        End If
'        strWHERE = strWHERE & " tProduct.P_Publisher = '" & pcolPublisher(i) & "'"
'    Next i
'
'    strSQL = "UPDATE tProduct SET tProduct.P_SupplierID = " & pNewTPID _
'                    & " WHERE (( " & strWHERE & " ) " _
'                    & " AND ((tProduct.P_SupplierID)= " & pLastTPID & ")); "
'    Me.RunSQL strSQL
'
'EXIT_Handler:
'    Set cmd = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.SupplierChange(pcolPublisher,pLastTPID,pNewTPID,pMsg,pConn)", _
'         Array(pcolPublisher, pLastTPID, pNewTPID, pMsg, pConn)
'End Function

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
'
'Private Sub Class_Initialize()
'    On Error GoTo errHandler
''    Set objOp = New a_Operation
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.Class_Initialize"
'End Sub

'Private Sub Class_Terminate()
'    On Error GoTo errHandler
'    Set objOp = Nothing
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.Class_Terminate"
'End Sub
'Public Sub MarkOrdersAsCancelled(pSince As Date)
'    On Error GoTo errHandler
'Dim oUtil As z_SQL
'Dim bAnswer As Boolean
'    bAnswer = True
'    If pSince > DateAdd("m", -6, Date) Then
'        RaiseEvent OrderStatus("No", bAnswer)
'    ElseIf pSince > DateAdd("yyyy", -1, Date) Then
'        RaiseEvent OrderStatus("Warning", bAnswer)
'    End If
'    If bAnswer <> False Then
'        RunProc "sp_MarkOldOrdersAsCancelled", Array(pSince), ""
'    End If
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.MarkOrdersAsCancelled(pSince)", pSince
'End Sub
'Function DropTable(pName, pMsg As String, Optional pConn As ADODB.Connection) As Long
'    On Error GoTo errHandler
'Dim cmd As ADODB.Command
'Dim strSQL As String
'
'    Set cmd = New ADODB.Command
'    If pConn Is Nothing Then
'        cmd.ActiveConnection = oPC.CO
'    Else
'        cmd.ActiveConnection = pConn
'    End If
'    cmd.CommandText = "DROP TABLE " & pName & ";"
'    cmd.CommandType = adCmdText
'    RaiseEvent Status(strMMsg & ": " & pMsg)
'    cmd.Execute
'
'EXIT_Handler:
'    DropTable = 0
'    Exit Function
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.DropTable(pName,pMsg,pConn)", Array(pName, pMsg, pConn)
'End Function
'Function RunProc(pName As String, pParams As Variant, pMsg As String, Optional pConn As ADODB.Connection) As Long
'    On Error GoTo errHandler
'Dim cmd As New ADODB.Command
'
'    Set cmd = New ADODB.Command
'    If pConn Is Nothing Then
'        cmd.ActiveConnection = oPC.CO
'    Else
'        cmd.ActiveConnection = pConn
'    End If
'    cmd.CommandText = pName
'    cmd.CommandType = adCmdStoredProc
'    RaiseEvent Status(strMMsg & ": " & pMsg)
'    If IsArray(pParams) Then
'        If UBound(pParams, 1) > -1 Then
'            retval = cmd.Execute(, pParams)
'        Else
'            cmd.Execute
'        End If
'    Else
'        cmd.Execute
'    End If
'
'    Set cmd = Nothing
'
'EXIT_Handler:
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.RunProc(pName,pParams,pMsg,pConn)", Array(pName, pParams, pMsg, pConn)
'End Function
'Function RunSQL(pSQL As String, Optional pConn As ADODB.Connection) As Long
'    On Error GoTo errHandler
'Dim cmd As New ADODB.Command
'Dim lngRecordsAffected As Long
'
'    Set cmd = New ADODB.Command
'    If pConn Is Nothing Then
'        cmd.ActiveConnection = oPC.CO
'    Else
'        cmd.ActiveConnection = pConn
'    End If
'    cmd.CommandText = pSQL
'    cmd.CommandType = adCmdText
'
'    cmd.Execute lngRecordsAffected
'    Set cmd = Nothing
'    RunSQL = lngRecordsAffected
'
'EXIT_Handler:
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.RunSQL(pSQL,pConn)", Array(pSQL, pConn)
'End Function
'Function RunGetRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset, Optional pConn As ADODB.Connection) As Long
'    On Error GoTo errHandler
'
'
'Dim cmd As New ADODB.Command
'Dim lngRecordsAffected As Long
'
'    Set cmd = New ADODB.Command
'    If pConn Is Nothing Then
'        cmd.ActiveConnection = oPC.CO
'    Else
'        cmd.ActiveConnection = pConn
'    End If
'    cmd.CommandText = pName
'    cmd.CommandType = pType
'    cmd.CommandTimeout = 8
'    RaiseEvent Status(strMMsg & ": " & pMsg)
'    If IsArray(pParams) Then
'        If UBound(pParams, 1) > -1 Then
'            Set rs = cmd.Execute(, pParams)
'        Else
'            Set rs = cmd.Execute()
'        End If
'    Else
'        Set rs = cmd.Execute()
'    End If
'    Set cmd = Nothing
'    RunGetRecordset = rs.RecordCount
'EXIT_Handler:
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.RunGetRecordset(pName,pType,pParams,pMsg,rs,pConn)", Array(pName, pType, pParams, _
'         pMsg, rs, pConn)
'End Function
'
'Function GetDynamicRecordset(pName As String, pType As enumCommandType, pParams As Variant, pMsg As String, rs As ADODB.Recordset, _
'                    Optional pConn As ADODB.Connection)
'    On Error GoTo errHandler
'    If pConn Is Nothing Then
'        rs.Open pName, oPC.CO, adOpenDynamic, adLockOptimistic
'    Else
'        rs.Open pName, pConn, adOpenDynamic, adLockOptimistic
'    End If
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.GetDynamicRecordset(pName,pType,pParams,pMsg,rs,pConn)", Array(pName, pType, _
'         pParams, pMsg, rs, pConn)
'End Function

'Function UpdateFromBookfind(Optional pStaffID As Long)
'    On Error GoTo errHandler
'Dim strArticle As String
'Dim strTitleNet As String
'
''     This is used when a new CD arrives.
''     It finds all records on tblProduct where there are no details except ISBN and looks on the CD
''     for details, if they are found then the record is updated.
'Dim dteStarted As Date
'    Dim rs As ADODB.Recordset
'    Dim oProdCode As New z_ProdCode
'    Dim oBF As a_BookFind
'    Dim strMsg As String
'    Dim StartTime
'    Dim X As Long
'    Dim iCancelled As Integer
'    Dim strShortname As String
'    Dim lngProgress As Long
'    Dim lngMax As Long
'    iCancelled = False
'
'        RaiseEvent Status("Selecting records to update . . .")
'        Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient
'        rs.Open "SELECT * FROM vBooksWithISBNs WHERE ISNULL(P_SKIPBFWASH,0) <> 1", oPC.CO, adOpenKeyset, adLockOptimistic
'        rs.MoveLast
'        lngProgress = 0
'        lngMax = rs.RecordCount
'        RaiseEvent Progress(lngProgress, lngMax)
'        dteStarted = Now()
'        RaiseEvent Status("Updating from CD . . ." & "Started:" & Format(dteStarted, "dd/mm/yyyy Hh:Nn") & ", Records: " & rs.RecordCount)
'        StartTime = Now
'        rs.MoveFirst
'        Set oBF = New a_BookFind
'    Do While Not rs.EOF
'        lngProgress = lngProgress + 1
'        RaiseEvent Progress(lngProgress, lngMax)
'        oProdCode.Load CStr(rs![P_Code])
'        If Not oProdCode.IsCode Then
'            GoTo ENDOFLOOP
'        End If
'        If oBF.FetchFromBF(CStr(rs![P_Code])) = 0 Then
'
'                rs![P_MainAuthor] = Left$(FNS(oBF.MainAuthor), rs.Fields("P_MainAuthor").DefinedSize)
'                rs![P_Code] = Left$(FNS(oBF.Code), rs.Fields("P_Code").DefinedSize)
'                StripArticle FNS(Left$(FNS(oBF.Title), rs.Fields("P_Title").DefinedSize)), strArticle, strTitleNet
'                rs!P_Title = strTitleNet
'                rs!P_Article = strArticle
'                rs![P_SubTitle] = Left$(FNS(oBF.SubTitle), rs.Fields("P_SubTitle").DefinedSize)
'                rs![P_STATUS] = Left(FNS(oBF.Availability), 1) ' rs.Fields("P_SubTitle").DefinedSize)
'                rs![P_Bindingcode] = Left$(FNS(oBF.BindingCode), rs.Fields("P_Bindingcode").DefinedSize)
'                rs![P_Edition] = Left$(FNS(oBF.Edition), rs.Fields("P_Edition").DefinedSize)
'                rs![P_Publisher] = Left$(FNS(oBF.PublisherName), rs.Fields("P_Publisher").DefinedSize)
'                rs![P_BFSupplierCode] = Left$(FNS(oBF.DistributorCode), rs.Fields("P_BFSupplierCode").DefinedSize)
'                strShortname = rs![P_Publisher]
'
'
''We are fleshing out the a_Product table, and when there is a skeleton record (Defined as having no Publisher
''name) then we fill in the other fields off Bookfind and then get the matching publisher record and see
''whether it is on the TP table (we match by Publisher name). If not, we add a publisher record with
''ISBNPrefix.
'   'Get BF Publisher Name and ISBNPrefix
'   'Try to match on papyrus TP_Shortname
'   'if not matched then
'   '   add new publisher record etc.
'   'end if
'   'assign new publisher record ID to Title
'StartSearch2:
''       If Not SeekRec("tblTradingPartner", "TP_Shortname", "=", strShortname) Then
''            RetVal = GetMatchingBFPublisher(Trim$(BFRec.k1), ISBNPrefix, strImprintName, strPublisherName, strBFCode) 'returns ISBNPrefix in second param
''            If RetVal = True Then  'found on BookFind''
''
'                                            'This publisher is not on database , so add a record to
'                                            'tblTradingpartner and tblTradingPartner_Role and tblSupplier_Imprint
''                 RetVal = AddAPublisher(strImprintName, strPublisherName, Trim$(ISBNPrefix), strBFCode, strShortname, 0)
''                 GoTo StartSearch2
''            End If
''        Else
''            rs![P_PublisherID] = gSeekRst![Trading_Partner_ID]
''        End If
'
'                If IsNumeric(oBF.UKPrice) Then
'                    rs![P_UKPrice] = CCur(oBF.UKPrice)
'                Else
'                    rs![P_UKPrice] = Null
'                End If
'                If IsNumeric(oBF.LocalPrice) Then
'                    If oBF.LocalPrice > 0 Then
'                        rs![P_RRP] = FNN(oBF.LocalPrice * oPC.Configuration.DefaultCurrency.Divisor)
'                    Else
'                        rs![P_RRP] = FNN(oBF.UKPrice * IIf(oPC.Configuration.Currencies.FindByDescription("POUND") Is Nothing, 10, oPC.Configuration.Currencies.FindByDescription("POUND")) * oPC.Configuration.DefaultCurrency.Divisor)
'                    End If
'                End If
'                rs![P_SeriesTitle] = Left$(FNS(oBF.SeriesTitle), rs.Fields("P_SeriesTitle").DefinedSize)
'
'                rs![P_Pubdate] = Left$(FNS(oBF.PublicationDate), rs.Fields("P_Pubdate").DefinedSize)
'
'                rs![P_BIC] = Left(oBF.BFClassification, rs![P_BIC].DefinedSize)
'                rs.Update
'            Else
'                If DateDiff("yyyy", FND(rs!P_LastDateDelivered), Date) > 2 Then 'mark books received last more than 2 years ago and not on Bookfind as Out of Print
'                    rs!P_STATUS = "O"
'                    rs.Update
'                End If
'            End If
'        'exit if operation cancelled
'ENDOFLOOP:
'        rs.MoveNext
'        DoEvents
'    Loop
'
'    rs.Close
'
'    oPC.CO.Execute "INSERT INTO tOPERATION (OP_STARTEDAT,OP_ENDEDAT,OP_NOMINALDATE,OP_TYPE,OP_STARTEDBYID,OP_RESULT) VALUES ('" & ReverseDateTime(dteStarted) & "','" & ReverseDateTime(Now) & "','" & ReverseDate(Date) & "',BookDataWash," & pStaffID & ",1)"
'    RaiseEvent Status("Updating from CD . . . complete.")
'    Exit Function
'
'ERRH::
'    MsgBox Error
'    Exit Function
'    Resume
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.UpdateFromBookfind"
'End Function
'
'Function createpublisherlist()
'    On Error GoTo errHandler
'Dim Nam
'Dim Code
'Dim extra As String
'Dim strShortname
'Dim iReccount As Long
'Dim iRecsFound As Long
'Dim oBF As a_BookFind
'Dim rs As ADODB.Recordset
'Dim i As Long
'Dim strSearchTabs As String
'Dim strRetrieveTabs As String
'Dim strTemp As String
'Dim strResults As String
'Dim strCode As String
'Dim iresultsFound As Variant
'    oPC.CO.Execute "DELETE FROM tDISTR"
'    Set oBF = New a_BookFind
'    oBF.InitializeScan "L2", "A", iresultsFound
'    oPC.BF.DefineOutputFormat ("L2")
'    i = 1
'    Do While i <= 999999
'        strCode = oBF.GetNextTerm
'        If strCode > "" Then
'            oPC.CO.Execute "INSERT INTO tDISTR (DISTR_ID) VALUES ('" & strCode & "')"
'            i = i + 1
'        Else
'            Exit Do
'        End If
'    Loop
'    Set rs = New ADODB.Recordset
'    rs.Open "SELECT * FROM tDISTR", oPC.CO
'    oPC.BF.DefineOutputFormat ("DN,AD,AP,TL,FX,EM")
'    Do While Not rs.EOF
'        strTemp = "FIND" & Chr(9) & "L2" & Chr(9) & FNS(rs!DISTR_ID)
'        Check oBF.Search(strTemp, iRecsFound) = 0, EXC_GENERAL, "Error in oBF:Search"
'        If iRecsFound > 0 Then
'            Check oBF.GetRecordByNum(1) = 0, EXC_GENERAL, "Error in oBF:GetRecordByNum"
'            oBF.LoadDistrRec (1)
'            oPC.CO.Execute "UPDATE tDISTR SET DISTR_NAME = '" & SQLQuotes(oBF.DistributorName) & "' WHERE DISTR_ID = '" & FNS(rs!DISTR_ID) & "'"
'        End If
'        rs.MoveNext
'    Loop
'    rs.Close
'    Set oBF = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.createpublisherlist"
'End Function
'Function bfPubSearch(fld As String, arg As String) As Long
'    On Error GoTo errHandler
'Dim iErr As Long
'Dim strSearchSpec As String
'Dim iRecsFound As Long
'    ChDrive "C:"
'    ChDir "C:\Bookfind"
'
'    strSearchSpec = "FIND" & Chr$(9) & fld & Chr$(9) & arg
'    Check oPC.BF.Search(strSearchSpec, iRecsFound), EXC_GENERAL, "Error in BF.Search"
'    bfPubSearch = iRecsFound
'    Exit Function
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Batch.bfPubSearch(fld,arg)", Array(fld, arg)
'End Function
'
