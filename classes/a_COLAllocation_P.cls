VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_COLAllocation_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GenerateCOLAllocationset(Optional pDELID As Long, Optional pType As String, Optional pOpCode As Long, _
                    Optional pCustID As Long, Optional pSupplierID As Long, Optional pFrom As String, Optional pTo As String, Optional bComplete As Boolean)
    On Error GoTo errHandler
Dim strSQL As String
Dim lngRecords As Long
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    oPC.COShort.BeginTrans
    If pDELID > 0 Then
    ' Remember: we must make sure we have not already allocated for this COL
        LoadOrderFulfilment_ByDELID pDELID
    Else
        If pOpCode > 0 Then
            LoadOrderFulfilment_ByStaffID pOpCode, bComplete
        ElseIf pCustID > 0 Then
            LoadOrderFulfilment_ByCustomerID pCustID, bComplete
        ElseIf pSupplierID > 0 Then
            LoadOrderFulfilment_BySupplierID pSupplierID, bComplete
        ElseIf pFrom > "" And pTo > "" Then
            LoadOrderFulfilment_ByName pFrom, pTo, bComplete
        End If
    End If

    oPC.COShort.CommitTrans
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    ErrPreserve
    If Error = "Timeout expired" Then
        MsgBox "This request has taken longer than allowed, please try to select fewer records. If it recurs, please call support.", vbInformation + vbOKOnly, "Status"
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.GenerateCOLAllocationset(pDELID,pType,pOpCode,pCustID,pFrom,pTo)", _
         Array(pDELID, pType, pOpCode, pCustID, pFrom, pTo), EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function
Private Sub LoadOrderFulfilment_ByName(pFrom As String, pTo As String, bComplete As Boolean)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strWSName As String

    strWSName = oPC.NameOfPC

    Set cmd = New ADODB.Command
    cmd.CommandText = "LoadOrderFulfilment_ByName"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@STARTNAME", adVarChar, adParamInput, 50, pFrom)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@ENDNAME", adVarChar, adParamInput, 50, pTo)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@WSNAME", adVarChar, adParamInput, 50, strWSName)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@COMPLETEONLY", adBoolean, adParamInput, , bComplete)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

    Exit Sub

    Exit Sub
errHandler:
    ErrPreserve
    If Error = "Timeout expired" Then
        MsgBox "This request has taken longer than allowed, please try to select fewer records. If it recurs, please call support.", vbInformation + vbOKOnly, "Status"
        Resume Next
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.LoadOrderFulfilment_ByName(pFrom,pTo)", Array(pFrom, pTo)
End Sub
Private Sub LoadOrderFulfilment_ByStaffID(pSTAFFID As Long, bComplete As Boolean)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strWSName As String

    strWSName = oPC.NameOfPC

    Set cmd = New ADODB.Command
    cmd.CommandText = "LoadOrderFulfilment_ByStaffID"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@STAFFID", adInteger, adParamInput, , pSTAFFID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@WSNAME", adVarChar, adParamInput, 50, strWSName)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@COMPLETEONLY", adBoolean, adParamInput, , bComplete)
    cmd.Parameters.Append par
    Set par = Nothing
    
   
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.LoadOrderFulfilment_ByStaffID(pStaffID,bComplete)", Array(pSTAFFID, _
         bComplete)
End Sub
Private Sub LoadOrderFulfilment_ByDELID(pDELID As Long)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strWSName As String

    strWSName = oPC.NameOfPC

    Set cmd = New ADODB.Command
    cmd.CommandText = "LoadOrderFulfilment_ByDELIDb"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@DELID", adInteger, adParamInput, , pDELID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

    Exit Sub

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_COLAllocation_P.LoadOrderFulfilment_ByStaffID(pStaffID,bComplete)", Array(pStaffID, _
'         bComplete)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.LoadOrderFulfilment_ByDELID(pDELID)", pDELID
End Sub

Private Sub LoadOrderFulfilment_BySupplierID(pSupplierID As Long, bComplete As Boolean)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strWSName As String

    strWSName = oPC.NameOfPC

    Set cmd = New ADODB.Command
    cmd.CommandText = "LoadOrderFulfilment_BySupplierID"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@SUPPLIERID", adInteger, adParamInput, , pSupplierID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@WSNAME", adVarChar, adParamInput, 50, strWSName)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@COMPLETEONLY", adBoolean, adParamInput, , bComplete)
    cmd.Parameters.Append par
    Set par = Nothing
    
   
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.LoadOrderFulfilment_BySupplierID(pSupplierID,bComplete)", _
         Array(pSupplierID, bComplete)
End Sub '
Private Sub LoadOrderFulfilment_ByCustomerID(pCustomerID As Long, bComplete As Boolean)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim strWSName As String

    strWSName = oPC.NameOfPC

    Set cmd = New ADODB.Command
    cmd.CommandText = "LoadOrderFulfilment_ByCustomerID"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 360
    
    Set par = cmd.CreateParameter("@CUSTOMERID", adInteger, adParamInput, , pCustomerID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@WSNAME", adVarChar, adParamInput, 50, oPC.WorkstationName)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@COMPLETEONLY", adBoolean, adParamInput, , bComplete)
    cmd.Parameters.Append par
    Set par = Nothing
    
   
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.LoadOrderFulfilment_ByCustomerID(pCustomerID,bComplete)", _
         Array(pCustomerID, bComplete)
End Sub '

Private Function CreateCOLAllocations(pDELID As Long) As Boolean
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CREATECOLALLOCATIONS"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@DELID", adInteger, adParamInput, , pDELID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    Set par = cmd.CreateParameter("@RESULT", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CreateCOLAllocations = (cmd.Parameters(1) = 0)
    Set cmd = Nothing

    Exit Function
End Function

Public Function Save(ByVal buffer As String, pType As String, InvoiceOrAppro As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As COLAllocationProps
Dim udtData As COLAllocationData
Dim oPB As PropertyBag
Dim oPBOUT As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim oOp As New a_Operation
Dim lngResult As Long
Dim oSM As z_StockManager

    Set oPB = New PropertyBag
    Set oPBOUT = New PropertyBag
    arBuffer = buffer
    oPB.Contents = arBuffer
    Set rs = New ADODB.Recordset
    
    oPC.COShort.BeginTrans
    If pType <> "DELIVERY" Then
        oOp.BeginEdit
        oOp.StartedAt = Now
        oOp.NominalDate = Date
        oOp.TypeID = AllocationStockToOrders
        oOp.ApplyEdit
        oOp.BeginEdit
    End If
    For lngIndex = 1 To oPB.ReadProperty("Count")
        udtData.buffer = oPB.ReadProperty("Item" & CStr(lngIndex))
        LSet udtProps = udtData

        If (Not udtProps.IsDeleted) And ((udtProps.AllocatedQty > 0) Or (udtProps.AllocatedQtySS > 0)) Then
            strSQL = "SELECT * FROM tCOLAlloc WHERE COLALLOC_ID = " & udtProps.ID
            rs.Open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
            If udtProps.IsNew Then rs.AddNew
            If Not rs.eof Then
                With rs
                    .Fields("COLALLOC_COL_ID") = udtProps.COLID
                    If oPC.AllowsSSInvoicing Then
                        .Fields("COLALLOC_Qty") = udtProps.AllocatedQty
                        .Fields("COLALLOC_QtySS") = udtProps.AllocatedQtySS
                    Else
                        .Fields("COLALLOC_Qty") = udtProps.AllocatedQty
                    End If
                    .Fields("COLALLOC_Status") = udtProps.Status
                    If pType <> "DELIVERY" Then
                        .Fields("COLALLOC_OP_ID") = oOp.ID
                    Else
                        .Fields("COLALLOC_OP_ID") = 0
                    End If
                    .Fields("COLALLOC_Action") = IIf(FNB(udtProps.ActionYN) = True, 1, 0)
                     .Update
                    udtProps.IsNew = False
                    udtProps.IsDirty = False
                End With
                LSet udtData = udtProps
                lngCount = lngCount + 1
                oPBOUT.WriteProperty "Item" & CStr(lngCount), udtData.buffer
                If pType = "DELIVERY" And udtProps.AllocatedQty > 0 Then
                    oPC.COShort.execute "UPDATE tPRODUCT SET P_QTYRESERVED = ISNULL(P_QTYRESERVED,0) + " & udtProps.AllocatedQty & ",P_QtyOnBackorder = dbo.NonNegative(ISNULL(P_QTYOnBackOrder,0) - " & udtProps.AllocatedQty & ") WHERE P_ID = '" & udtProps.PID & "'"
                    oPC.COShort.execute "UPDATE tSTOREP SET STP_QTYONBACKORDER = dbo.Nonnegative(ISNULL(STP_QTYONBACKORDER,0) - " & udtProps.AllocatedQty & "),STP_QTYRESERVED = ISNULL(STP_QTYRESERVED,0) + " & udtProps.AllocatedQty & " WHERE STP_P_ID = '" & udtProps.PID & "' AND STP_ST_ID = " & oPC.Configuration.DefaultStoreID
                    oPC.COShort.execute "UPDATE tCOL SET COL_QTYDispatched = ISNULL(COL_QtyDispatched,0) + " & udtProps.AllocatedQty & " WHERE COL_ID = " & udtProps.COLID, lngResult
                    oPC.COShort.execute "UPDATE tCOL SET COL_Fulfilled = 'FUL' WHERE COL_QTY - COL_QtyDispatched <= 0 AND COL_ID = " & udtProps.COLID, lngResult
                    strSQL = "UPDATE tTR Set TR_STATUS = 4 WHERE TR_ID IN " _
                                & " (SELECT TR_ID  FROM tCOL JOIN tTR ON tCOL.COL_TR_ID = tTR.TR_ID WHERE tCOL.COL_TR_ID = " & rs.Fields("COLALLOC_CO_ID") _
                                & " GROUP BY TR_CODE ,TR_ID " _
                                & " Having (MAX(COL_Fulfilled) = 'CAN') or (MAX(COL_Fulfilled) = 'FUL'))"
                    oPC.COShort.execute strSQL, lngResult
                End If
            End If
            rs.Close
        Else
          DeleteCOLALLOC udtProps.ID, pType
    End If
    Next
    oPBOUT.WriteProperty "Count", lngCount
    
    Set oPB = Nothing
    Set rs = Nothing
    
    Save = oPBOUT.Contents
    Set oPBOUT = Nothing
    If pType = "DELIVERY" Then
        oPC.COShort.execute "INSERT INTO tRM(RM_COL_ID,RM_P_ID,RM_Date,RM_Qty,RM_TP_ID,RM_DELL_ID,RM_USER_ID,RM_TYPE) " _
                    & " SELECT COL_ID,COL_P_ID,GETDATE(),COLALLOC_QTY,TR_TP_ID,COLALLOC_DEL_ID,0,'I'  " _
                    & " FROM tCOLALLOC JOIN tTR ON COLALLOC_CO_ID = TR_ID JOIN tCOL ON COLALLOC_COL_ID = COL_ID " _
                    & " WHERE COLALLOC_QTY > 0 AND COLALLOC_COL_ID = " & udtProps.COLID, lngResult
    Else
        oOp.Endedat = Now
        oOp.ApplyEdit
        Set oOp = Nothing
    End If
    oPC.COShort.CommitTrans

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.Save(buffer,pType)", Array(buffer, pType), EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Function
Public Function Fetch(Optional TRID As Long, Optional bReview As Boolean, Optional WSName As String) As String
    On Error GoTo errHandler
Dim lngCount As Long
Dim udtProps As COLAllocationProps
Dim udtData As COLAllocationData
Dim pblnNoRecsReturned As Boolean
Dim rs As ADODB.Recordset
Dim dPOL As d_POLine
Dim lngID As Long
Dim objPB As PropertyBag
Dim strSQL As String
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    Set rs = New ADODB.Recordset
    If TRID > 0 Then
        If bReview Then
            strSQL = "SELECT * FROM COLsOS  WHERE COLALLOC_DEL_ID = " & TRID & " ORDER By P_TITLE,P_CODE,TRDATE"
        Else
            strSQL = "SELECT * FROM COLsOS  WHERE COALESCE(COL_Fulfilled,'OS') = 'OS'  and COLALLOC_DEL_ID = " & TRID & " ORDER By P_TITLE,P_CODE,TRDATE"
        End If
    Else
        If bReview Then
            strSQL = "SELECT * FROM COLsOS WHERE COLALLOC_DEL_ID = 0 ORDER By P_TITLE,P_CODE,TRDATE"
        Else
            strSQL = "SELECT * FROM COLsOS WHERE COALESCE(COL_Fulfilled,'OS') = 'OS' and ISNULL(COLALLOC_DEL_ID,0) = 0  AND COLALLOC_WSNAME = '" & WSName & "' ORDER By P_TITLE,P_CODE,TRDATE"
        End If
    End If
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open strSQL, oPC.COShort, adOpenForwardOnly
    If rs.eof Then
        pblnNoRecsReturned = True
        objPB.WriteProperty "Count", 0
        Fetch = objPB.Contents
        Set objPB = Nothing
        Exit Function
    End If
    
    On Err GoTo errHandler
    Do While Not rs.eof
        With udtProps
            udtProps.ID = FNN(rs!COLALLOC_ID)
            udtProps.Ref = FNS(rs!COL_Ref)
            udtProps.QuotedPrice = FNN(rs!COL_PRICE)
            If oPC.AllowsSSInvoicing Then
                udtProps.OrderedQty = FNN(rs!COL_QtyFirm)
                udtProps.OrderedSSQTY = FNN(rs!COL_QtySS)
            Else
                udtProps.OrderedQty = FNN(rs!COL_Qty)
            End If
            udtProps.DeliveredSoFar = FNN(rs!COL_QtyDispatched)
            udtProps.DeliveredSoFarSS = FNN(rs!DispatchedQtySS)
            udtProps.DeliveredSoFarFirm = FNN(rs!DispatchedQtyFirm)
            udtProps.QtyOnHand = FNN(rs!COLALLOC_QtyOHwithsubs) ' FNN(rs!P_QtyOnHand)
            udtProps.QtyReserved = FNN(rs!COLALLOC_Qty)
            udtProps.QtyonCO = FNN(rs!P_QtyOnBackorder)
            udtProps.QtyonOrder = FNN(rs!P_QtyOnOrder)
            udtProps.COLID = FNN(rs!COL_ID)
            udtProps.CustomerName = FNS(rs!TP_Name)
            udtProps.CustomerInitials = FNS(rs!TP_Initials)
            udtProps.CustomerTitle = FNS(rs!TP_Title)
            udtProps.CustomerAcno = FNS(rs!TP_ACno)
            udtProps.DepositValue = (FNDBL(rs!COL_Deposit) - FNDBL(rs!COL_DepositRefunded)) / oPC.Configuration.DefaultCurrency.Divisor
            udtProps.PID = FNS(rs!P_ID)
            udtProps.OrderCode = FNS(rs!TRCODE)
            udtProps.Title = FNS(rs!P_Title)
            udtProps.Phone = FNS(rs!ADD_Phone)
            udtProps.code = FNS(rs!CodeF)
            udtProps.OrderDate = FND(rs!TRDATE)
            udtProps.AllocatedQty = FNN(rs!COLALLOC_Qty)
            udtProps.UsesSubstitutesYN = FNS(rs!COLALLOC_SUBSYN)
            udtProps.ProductOH = FNN(rs!COLALLOC_QtyOH)
            udtProps.ProductRES = FNN(rs!COLALLOC_QtyRES)
            udtProps.WSLock = FNS(rs!COLALLOC_LOCKEDBY)
            udtProps.Note = FNS(rs!COLALLOC_NOTE)
            udtProps.CustomerBlocked = FNB(rs!TP_Blocked)
            udtProps.CreateInvoice = FNB(rs!COLALLOC_CreateInvoice)
            udtProps.Status = 1
            
            udtProps.IsNew = False
            udtProps.IsDirty = False
            udtProps.IsDeleted = False
            
            LSet udtData = udtProps
            lngCount = lngCount + 1
            objPB.WriteProperty "Item" & CStr(lngCount), udtData.buffer
            rs.MoveNext
        End With
    Loop
    With objPB
          .WriteProperty "Count", lngCount
          Fetch = .Contents
    End With
    Set objPB = Nothing
    
EXIT_Handler:
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.Fetch(TRID,bReview)", Array(TRID, bReview)
End Function
Private Sub DeleteCOLALLOC(COLAllocationID As Long, pType As String)
    On Error GoTo errHandler
Dim strSQL As String
    If pType = "DELIVERY" Then
        strSQL = "DELETE FROM tCOLAlloc WHERE COLALLOC_ID=" & COLAllocationID
    Else
        strSQL = "DELETE FROM tCOLAllocation WHERE COLALLOC_ID=" & COLAllocationID
    End If
    oPC.COShort.execute strSQL
'H:
'  oor.Setor , "a_COLAllocation:DeleteCOLALLOC:" & or, Date, "", "", ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.DeleteCOLALLOC(COLAllocationID,pType)", Array(COLAllocationID, pType), EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Sub

Public Sub DeleteAllforDEL(pDELID As Long, pType As String)
    On Error GoTo errHandler
Dim strSQL As String
    If pType = "DELIVERY" Then
        strSQL = "DELETE FROM tCOLAlloc JOIN tDELL ON DELL_ID = COLALLOC_DELL_ID WHERE DELL_TR_ID = " & pDELID
    Else
        strSQL = "DELETE FROM tCOLAllocation JOIN tDELL ON DELL_ID = COLALLOC_DELL_ID WHERE DELL_TR_ID = " & pDELID
    End If
    oPC.COShort.execute strSQL
'H:
'  oor.Setor , "a_COLAllocation:DeleteAllforDEL:" & or, Date, "", "", ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.DeleteAllforDEL(pDELID,pType)", Array(pDELID, pType), EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Sub


Public Function GenerateCOLAllocationset_TFR(Optional pTFRID As Long)
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------


    Set cmd = New ADODB.Command
    cmd.CommandText = "INSERT_COLALLOC_TFR"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@TFRID", adInteger, adParamInput, , pTFRID)
    cmd.Parameters.Append par
    Set par = Nothing
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    Set cmd = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_COLAllocation_P.GenerateCOLAllocationset_TFR(pTFRID)", pTFRID
End Function

Public Function DeleteAllocatedLocks(WSName As String)
 Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
   
    oPC.COShort.execute "DELETE FROM tCOLALLOC WHERE COLALLOC_WSNAME = '" & WSName & "'"
    oPC.COShort.execute "UPDATE tCOL SET COL_LOCK = '' WHERE COL_LOCK = '" & WSName & "'"
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
  
End Function

