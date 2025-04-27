VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Invoice_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As InvoiceProps
Dim udtData As InvoiceData
Dim objPersist As a_InvoiceLine_P
Dim objPB As PropertyBag
Dim oUtil As z_SQL
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    Set oUtil = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngResult = oUtil.RunGetRecordset("SELECT * FROM Invoice_Fetch WHERE TR_ID = " & TRID, enText, Array(), "", rs)
    
    
    If Not rs.eof Then
        With udtProps
            .InvoiceID = rs!TR_ID
        '    .ID = rs!TR_GlobalID
            .TPID = FNN(rs!TR_TP_ID)
            .BillToAddressID = FNN(rs.Fields("AddressDOCID"))
            .DelToAddressID = FNN(rs.Fields("AddressGOODSID"))
            .CurrencyID_Foreign = FNN(rs!I_Curr_ID)
            .Status = FNStatus(rs!TR_Status)
            If .CurrencyID_Foreign > 0 And .Status = 2 Then
                    .CurrencyFactor = oPC.Configuration.Currencies.FindCurrencyByID(.CurrencyID_Foreign).Factor
            Else
                    .CurrencyFactor = FNDBL(rs!I_CUrrRate)
            End If
            .VATRate = FNDBL(rs!I_VATRate)
            .DiscountRate = FNDBL(rs!I_DiscountRate)
            .TotalDiscount = FNN(rs!I_DiscountAmount)
            .TotalVAT = FNN(rs!I_VATAmount)
            .ForAttn = FNS(rs!I_ForAttn)
            .TotalPayable = FNN(rs!I_PayableAmount)
            .COMPID = FNN(rs!TR_COMP_ID)
            .ShowVAT = FNB(rs!I_ShowVAT)
            .CurrencyFormat = FNS(rs!CURR_Format)
            .BusPhone = FNS(rs!ADD_BusPhone)
            .TPACCNum = FNS(rs!TP_ACno)
            .Log = Trim(IIf(IsNull(rs!TR_Log), "", rs!TR_Log))
            .TPFax = FNS(rs!ADD_Fax)
            .TPPhone = FNS(rs!TP_Phone)
            .TPNAME = FNS(rs!TP_Name)
            .Memo = FNS(rs!TR_NOTE)
            .DOCCode = FNS(rs!TR_CODE)
            .Proforma = FNB(rs!I_Proforma)
            .DOCDate = FND(rs!TR_DATE)
            .ProcessingDate = FND(rs!TR_ProcessingDATE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .StaffID = FNN(rs!TR_StaffID)
            .IsPreDelivery = IIf(IsNull((rs!I_IsPredelivery)) = True, False, FNB((rs!I_IsPredelivery)))
            If .StaffID > 0 Then
                .StaffName = oPC.Configuration.Staff.FindStaffByID(.StaffID).Shortname
                .Signature = oPC.Configuration.Staff.FindStaffByID(.StaffID).Signature
                .StaffEmail = oPC.Configuration.Staff.FindStaffByID(.StaffID).EMail
            End If
        '    .ExchangeID = FNS(rs!Exchange)
            .SalesRepID = FNN(rs!I_REPID)
            .SalesRepName = FNS(rs!SalesRepName)
            .CustPaid = FNB(rs!I_InvoicePaid)
            .CommPaid = FNB(rs!I_CommissionPaid)
            .VATable = FNB(rs!I_Vatable)
            .ShowVAT = FNB(rs!I_ShowVAT)
            .NonVATDocument = FNB(rs!TR_NonVATDocument)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            Set oUtil = Nothing
            LSet udtData = udtProps
            Set objPersist = New a_InvoiceLine_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "ILS", objPersist.Fetch(udtProps.InvoiceID, "I")
              Fetch = .Contents
            End With
            Set objPB = Nothing
            Set objPersist = Nothing
        End With
    Else
        ' force an or
       ' rs.MoveNext
    End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXITH:
  Exit Function
  
errHandler:
    
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    RlsObjs oPC.COShort
    ErrorIn "a_Invoice_P.Fetch(TRID)", TRID
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim bUpdateStock As Boolean
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As InvoiceProps
Dim udtData As InvoiceData
Dim objPersist As a_InvoiceLine_P
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

    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.InvoiceID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
  '  On Error GoTo H
    With udtProps
        If .IsNew Then
            rs.AddNew
           ' If .DOCDate = CDate(0) Then .DOCDate = .CaptureDate
        End If
        rs!TR_Type = 3
        rs!TR_TP_ID = FNN(.TPID)
        rs!TR_COMP_ID = FNN(.COMPID)
        rs!TR_DATE = FND(.DOCDate)
        rs!TR_CaptureDate = FND(.CaptureDate)
        rs!TR_CODE = FNS(.DOCCode)
        rs!TR_NOTE = FNS(.Memo)
        rs!TR_AddressDOCID = FNN(.BillToAddressID)
        rs!TR_AddressGoodsID = FNN(.DelToAddressID)
        rs!TR_Status = FNN(.Status)
        rs!TR_StaffID = FNN(.StaffID)
        rs!TR_NonVATDocument = FNB(.NonVATDocument)

        rs.Update
        If .IsNew Then
           rs.MoveLast
           .InvoiceID = rs!TR_ID
        End If
        rs.Close
        
    End With

    Set rs = Nothing
    
    Set rs = New ADODB.Recordset
    strSQL = "SELECT * FROM tInvoice WHERE I_ID =" & udtProps.InvoiceID
    rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs!I_ID = FNN(.InvoiceID)
        rs!I_Curr_ID = FNN(.CurrencyID_Foreign)
        rs!I_Proforma = FNB(.Proforma)
        rs!I_REPID = FNN(.SalesRepID)
        rs!I_ForAttn = FNS(.ForAttn)
        rs!I_CUrrRate = FNDBL(.CurrencyFactor)  '   Conversion rate
        rs!I_DiscountRate = FNDBL(.DiscountRate)  '   Invoice discount (general) rate
        rs!I_VATRate = FNDBL(.VATRate)  '
        rs!I_Vatable = FNB(.VATable)
        rs!I_DiscountAmount = FNDBL(.TotalDiscount) / oPC.Configuration.DefaultCurrency.Divisor
        rs!I_VATAmount = FNDBL(.TotalVAT) / oPC.Configuration.DefaultCurrency.Divisor
        rs!I_PayableAmount = FNDBL(.TotalPayable) / oPC.Configuration.DefaultCurrency.Divisor
        rs!I_ShowVAT = FNB(.ShowVAT)
        rs!I_IsPredelivery = FNB(.IsPreDelivery)

        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    
    If (Not oPC.GetProperty("AllowsGDNs") = "TRUE" Or FNB(udtProps.IsPreDelivery) = False) And (udtProps.Proforma = False) Then
        Set rs = New ADODB.Recordset
        strSQL = "SELECT * FROM tDN WHERE DN_ID =" & udtProps.InvoiceID
        rs.Open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
        
        With udtProps
            If rs.RecordCount = 0 Then rs.AddNew
            rs!DN_ID = FNN(.InvoiceID)
            rs!DN_I_ID = FNN(.InvoiceID)
            rs.Update
        End With
        rs.Close
        Set rs = Nothing
    End If
    
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_InvoiceLine_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "ILS", objPersist.Save(objPB.ReadProperty("ILS"), udtProps.InvoiceID, "I", FNB(udtProps.IsPreDelivery))
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
    
      Save = objPBOut.Contents
      Set objPBOut = Nothing
      
    If Not udtProps.Proforma Then
        oPC.COShort.execute "Execute dbo.AllocateSalesCommission " & udtProps.InvoiceID & "," & FNN(udtProps.SalesRepID)
    End If
  
EXITH:
    Exit Function
errHandler:
    If Error = "Timeout expired" Then
        LogSaveToFile "so-called TimeoutExpired"
        Save = "TIMEOUT"
        GoTo EXITH
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    RlsObjs oPC.COShort
    ErrorIn "a_Invoice_P.Save(buffer)", buffer, , rs
End Function


Private Sub DeleteInvoice(ID As Long)
    On Error GoTo errHandler
    oPC.COShort.execute "DELETE FROM tPI WHERE PI_ID=" & ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice_P.DeleteInvoice(ID)", ID
End Sub

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_InvoiceLine_P
  
    oPC.COShort.BeginTrans
    oPC.COShort.execute "DELETE FROM tInvoice WHERE ID=" & ID
  
    Set objPersist = New a_InvoiceLine_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.CommitTrans
    Exit Sub
  
'H:
'    oPC.COShort.RollbackTrans
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Invoice_P.DeleteObject(ID)", ID
End Sub

