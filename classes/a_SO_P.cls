VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_PO_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As POProps
Dim udtData As POData
Dim objPersist As a_POL_P
Dim objPB As PropertyBag
Dim oUtil As z_SQL
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    Set oUtil = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngResult = oUtil.RunGetRecordset("SELECT * FROM PO_Fetch WHERE PO_ID = " & TRID, enText, Array(), "", rs)
    
    
    If Not rs.eof Then
        With udtProps
            .TRID = rs!TR_ID
            .TPID = rs!TR_TP_ID
            .DocStoreID = FNN(rs.fields("AddressDOCID"))
            .DELTOStoreID = FNN(rs.fields("AddressGOODSID"))
            If .DELTOStoreID = 0 Then
                .DELTOStoreID = .DocStoreID
            End If
            If oPC.Configuration.Stores.FindStoreByID(.DELTOStoreID) Is Nothing Then
                .DelToAddress = oPC.Configuration.DefaultStore.DelAddress
            Else
                .DelToAddress = oPC.Configuration.Stores.FindStoreByID(.DELTOStoreID).DelAddress
            End If
            .TotalVAT = FNN(rs!PO_VATAmount)
            .TotalPayable = FNN(rs!PO_PayableAmount)   'Always store the local amount
            .COMPID = FNN(rs!TR_COMP_ID)
            .TPPhone = FNS(rs!TP_Phone)
            .TPPhone2 = FNS(rs!ADD_BusPhone)
            .TPACCNum = FNS(rs!TP_ACno)
            .OrderType = FNS(rs!PO_OrderType)
            .TPFax = FNS(rs!ADD_Fax)
            .TPNAME = FNS(rs!TP_Name)
            .Memo = FNS(rs!TR_NOTE)
            .Log = Trim(IIf(IsNull(rs!TR_Log), "", rs!TR_Log))
            .CurrencyID = FNN(rs!PO_CURR_ID)
            .CurrencyRate = FNDBL(rs!PO_CURRRATE)
            .DispatchMethodID = FNN(rs!PO_DispatchModeID)
            .DOCCode = FNS(rs!TR_CODE)
            .TempETA = FND(rs!PO_ETA)

            If rs!TR_DATE > DateSerial(1995, 1, 1) Then .DOCDate = DateSerial(Year(rs!TR_DATE), Month(rs!TR_DATE), Day(rs!TR_DATE))
            .CaptureDate = FND(rs!TR_CaptureDate)
            .ProcessingDate = FND(rs!TR_ProcessingDATE)
            .StaffID = FNN(rs!TR_StaffID)
            If .StaffID > 0 Then
                If Not oPC.Configuration.Staff.FindStaffByID(.StaffID) Is Nothing Then
                    .StaffName = oPC.Configuration.Staff.FindStaffByID(.StaffID).Shortname
                    .Signature = oPC.Configuration.Staff.FindStaffByID(.StaffID).Signature
                    .StaffEmail = oPC.Configuration.Staff.FindStaffByID(.StaffID).EMail
                End If
            End If
            .Status = FNStatus(rs!TR_Status)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            Set oUtil = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_POL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "POLS", objPersist.Fetch(udtProps.TRID, udtProps.Status)
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
    ErrorIn "a_PO_P.Fetch(TRID)", TRID
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim OpenResult As Integer
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As POProps
Dim udtData As POData
Dim objPersist As a_POL_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte

    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData


'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.TRID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    With udtProps
        If .IsNew Then
            rs.AddNew
        End If
        rs!TR_Type = 5
        rs!TR_TP_ID = FNN(.TPID)
        rs!TR_DATE = FNDateSerial(.DOCDate)
        rs!TR_CaptureDate = FND(.CaptureDate)
       ' rs!TR_ProcessingDATE = FND(.ProcessingDate)
        rs!TR_CODE = FNS(.DOCCode)
        rs!TR_NOTE = FNS(.Memo)
        rs!TR_AddressDOCID = FNN(.DocStoreID)
        rs!TR_AddressGoodsID = FNN(.DELTOStoreID)
        rs!TR_Status = FNN(.Status)
        rs!TR_COMP_ID = oPC.Configuration.DefaultCOMPID
        rs!TR_StaffID = FNN(.StaffID)
        rs.Update
        If .IsNew Then
           rs.MoveLast
           .TRID = rs!TR_ID
        End If
        rs.Close
    End With

    strSQL = "SELECT * FROM tPO WHERE PO_ID =" & udtProps.TRID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then
            rs.AddNew
        End If
        rs!PO_ID = .TRID
        rs!PO_CURRRATE = FNN(.CurrencyRate)         'Conversion rate
        rs!PO_CURR_ID = FNN(.CurrencyID)
        rs!PO_VATAmount = FNN(.TotalVAT)
        rs!PO_PayableAmount = FNN(.TotalPayable)    'Always store the local amount
        rs!PO_Discount = FNN(.TotalDiscount)    'Always store the local amount
        rs!PO_OrderType = FNS(.OrderType)
        rs!PO_DispatchModeID = FNN(.DispatchMethodID)
        rs!PO_ETA = FND(.TempETA)

        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_POL_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "POLS", objPersist.Save(objPB.ReadProperty("POLS"), udtProps.TRID, FNS(udtProps.DOCCode))
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
    
      Save = objPBOut.Contents
      Set objPBOut = Nothing
      
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO_P.Save(buffer)", buffer, , rs
End Function


'Private Sub DeleteInvoice(ID As Long)
'    oPC.COShort.Execute "DELETE FROM tPI WHERE PI_ID=" & ID
'
'End Sub

Public Sub DeletePO(ByVal TRID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_POL_P
  
    oPC.COShort.BeginTrans
  
    Set objPersist = New a_POL_P
    objPersist.DeletePOLSet TRID
    Set objPersist = Nothing
    oPC.COShort.execute "DELETE FROM tTR WHERE TR_ID=" & TRID
    oPC.COShort.execute "DELETE FROM tPO WHERE PO_ID=" & TRID
    oPC.COShort.CommitTrans
    Exit Sub
  
'H:
'    oPC.COShort.RollbackTrans
'    Err.Raise err.Number
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_PO_P.DeletePO(TRID)", TRID, EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
End Sub


