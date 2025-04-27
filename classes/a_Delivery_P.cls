VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Delivery_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fetch(TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As DELProps
Dim udtData As DELData
Dim objPersist As a_DeliveryLine_P
Dim objPB As PropertyBag
Dim oUtil As z_SQL
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    Set oUtil = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngResult = oUtil.RunGetRecordset("SELECT * FROM DEL_Fetch WHERE TR_ID = " & TRID, enText, Array(), "", rs)
    
    
    If Not rs.eof Then
        With udtProps
            .TRID = FNN(rs!TR_ID)
            .TPID = FNN(rs!TR_TP_ID)
            .TPACCNum = FNS(rs!TP_ACno)
            .TPNAME = FNS(rs!TP_Name)
            .DOCCode = FNS(rs!TR_CODE)
            .Memo = FNS(rs!TR_NOTE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .ProcessingDate = FND(rs!TR_ProcessingDATE)
            .SupplierInvoiceDate = FND(rs!DEL_SUpplierInvoiceDate)
            .SupplierInvoiceRef = FNS(rs!DEL_SUpplierInvoiceRef)
            .CurrencyID = FNN(rs!DEL_CURR_ID)
            .CurrencyRate = FNDBL(rs!DEL_CURRCONVERSIONRATE)
        
            .BatchTotal = FNN(rs!DEL_BatchTotal)
            .BatchQtyTotal = FNN(rs!DEL_BatchQtyTotal)
            .BatchTotalExtras = FNN(rs!DEL_BatchTotalExtras)
            If rs!TR_DATE > DateSerial(1995, 1, 1) Then .DOCDate = DateSerial(Year(rs!TR_DATE), Month(rs!TR_DATE), Day(rs!TR_DATE))
            .Status = FNStatus(rs!TR_Status)
            .StaffID = FNN(rs!TR_StaffID)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            Set oUtil = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_DeliveryLine_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "DELLS", objPersist.Fetch(udtProps.TRID)
              Fetch = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
            End With
                       
    Else
        ' force an or
        rs.MoveNext
        
    End If
    
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery_P.Fetch(TRID)", TRID
End Function
Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As DELProps
Dim udtData As DELData
Dim objPersist As a_DeliveryLine_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim OpenResult As Integer

    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.TRID
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
  '  On Error GoTo H
    With udtProps
        If .IsNew Then rs.AddNew
        rs!TR_Type = 4
        rs!TR_TP_ID = FNN(.TPID)
        rs!TR_DATE = FNDateSerial(.DOCDate)
        rs!TR_CaptureDate = FND(.CaptureDate)
        rs!TR_CODE = FNS(.DOCCode)
        rs!TR_Status = .Status
        rs!TR_NOTE = FNS(.Memo)
        rs!TR_COMP_ID = .COMPID
        rs!TR_StaffID = .StaffID
        rs!TR_SentToCentralAt = Null
        rs.Update
        If .IsNew Then
           rs.MoveLast
           .TRID = rs!TR_ID
        End If
        rs.Close
    End With

    strSQL = "SELECT * FROM tDEL WHERE DEL_ID =" & udtProps.TRID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs!DEL_ID = .TRID
        rs!DEL_SUpplierInvoiceRef = FNS(.SupplierInvoiceRef)
        rs!DEL_SUpplierInvoiceDate = FND(.SupplierInvoiceDate)
        rs!DEL_CURR_ID = FNN(.CurrencyID)
        rs!DEL_BatchTotal = FNN(.BatchTotal)
        rs!DEL_BatchQtyTotal = FNN(.BatchQtyTotal)
        rs!DEL_BatchTotalExtras = FNN(.BatchTotalExtras)
        rs!DEL_CURRCONVERSIONRATE = oPC.Configuration.Currencies.FindCurrencyByID(FNN(.CurrencyID)).Factor
        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    If FNN(udtProps.BatchTotalExtras) > 0 Then
        strSQL = "SELECT * FROM tExtraCharges WHERE EXC_SOURCETR_ID =" & udtProps.TRID
        Set rs = New ADODB.Recordset
        rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
        With udtProps
            If rs.eof Then rs.AddNew
            rs!EXC_SourceTR_ID = .TRID
            rs!EXC_Amount = FNN(.BatchTotalExtras)
            rs.Update
        End With
        rs.Close
        Set rs = Nothing
    End If
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_DeliveryLine_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "DELLS", objPersist.Save(objPB.ReadProperty("DELLS"), udtProps.TRID)
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
    
      Save = objPBOut.Contents
      Set objPBOut = Nothing
    Dim oSQL As New z_SQL
   ' oSQL.RunProc "cmdSendDeliveriesToCentral", Array(), "Sending deliveries to central via service broker"
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery_P.Save(buffer)", buffer
End Function


Private Sub DeleteDELL(ID As Long)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tDELL WHERE DELL_ID=" & ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery_P.DeleteDELL(ID)", ID
End Sub

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_DeliveryLine_P
  
    oPC.COShort.BeginTrans
  
    Set objPersist = New a_DeliveryLine_P
    objPersist.DeleteDELL ID
    Set objPersist = Nothing
    oPC.COShort.execute "DELETE FROM tDEL WHERE DEL_ID=" & ID
    oPC.COShort.CommitTrans
    Exit Sub
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Delivery_P.DeleteObject(ID)", ID
End Sub


