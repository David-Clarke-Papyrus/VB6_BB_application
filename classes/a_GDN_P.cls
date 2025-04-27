VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_GDN_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As GDNProps
Dim udtData As GDNData
Dim objPersist As a_InvoiceLine_P
Dim objPB As PropertyBag
Dim oUtil As z_SQL
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    Set oUtil = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngResult = oUtil.RunGetRecordset("SELECT * FROM GDN_Fetch WHERE TR_ID = " & TRID, enText, Array(), "", rs)
    
    
    If Not rs.eof Then
        With udtProps
            .GDNID = rs!TR_ID
            .TPID = FNN(rs!TR_TP_ID)
            .BillToAddressID = FNN(rs.fields("AddressDOCID"))
            .DelToAddressID = FNN(rs.fields("AddressGOODSID"))
            .CurrencyID_Foreign = FNN(rs!TP_CURR_ID)
            .Status = FNStatus(rs!TR_Status)
'            If .CurrencyID_Foreign > 0 And .STATUS = 2 Then
'                    .CurrencyFactor = oPC.Configuration.Currencies.FindCurrencyByID(.CurrencyID_Foreign).Factor
'            Else
'                    .CurrencyFactor = FNDBL(rs!I_CUrrRate)
'            End If
'            .VATRate = FNDBL(rs!I_VATRate)
'            .DiscountRate = FNDBL(rs!I_DiscountRate)
'            .TotalDiscount = FNN(rs!I_DiscountAmount)
'            .TotalVAT = FNN(rs!I_VATAmount)
'            .ForAttn = FNS(rs!I_ForAttn)
'            .TotalPayable = FNN(rs!I_PayableAmount)
            .COMPID = FNN(rs!TR_COMP_ID)
            .CurrencyFormat = FNS(rs!CustomerCurrFormat)
            .BusPhone = FNS(rs!ADD_BusPhone)
            .TPACCNum = FNS(rs!TP_ACno)
            .Log = Trim(IIf(IsNull(rs!TR_Log), "", rs!TR_Log))
            .TPFax = FNS(rs!ADD_Fax)
            .TPPhone = FNS(rs!TP_Phone)
            .TPNAME = FNS(rs!TP_Name)
            .Memo = FNS(rs!TR_NOTE)
            .DOCCode = FNS(rs!TR_CODE)
            .DOCDate = FND(rs!TR_DATE)
            .ProcessingDate = FND(rs!TR_ProcessingDATE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .StaffID = FNN(rs!TR_StaffID)
            .InvoiceCode = FNS(rs!InvoiceCode)
            If .StaffID > 0 Then
                .StaffName = oPC.Configuration.Staff.FindStaffByID(.StaffID).Shortname
                .Signature = oPC.Configuration.Staff.FindStaffByID(.StaffID).Signature
                .StaffEmail = oPC.Configuration.Staff.FindStaffByID(.StaffID).EMail
            End If
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
              .WriteProperty "ILS", objPersist.Fetch(udtProps.GDNID, "G")
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
Dim udtProps As GDNProps
Dim udtData As GDNData
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
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.GDNID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
  '  On Error GoTo H
    With udtProps
        If .IsNew Then
            rs.AddNew
           ' If .DOCDate = CDate(0) Then .DOCDate = .CaptureDate
        End If
        rs!TR_Type = 13
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
           .GDNID = rs!TR_ID
        End If
        rs.Close
'        If FNN(.Status) = 3 Or FNN(.Status) = 4 Or FNN(.Status) = 5 Then
'            bUpdateStock = True
'        Else
'            bUpdateStock = False
'        End If
        
    End With

    strSQL = "SELECT * FROM tDN WHERE DN_ID =" & udtProps.GDNID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs!DN_ID = FNN(.GDNID)
        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_InvoiceLine_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "ILS", objPersist.Save(objPB.ReadProperty("ILS"), udtProps.GDNID, "G", False)
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
    
      Save = objPBOut.Contents
      Set objPBOut = Nothing
      
  
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
  Dim objPersist As a_GDNLine_P
  
    oPC.COShort.BeginTrans
    oPC.COShort.execute "DELETE FROM tInvoice WHERE ID=" & ID
  
    Set objPersist = New a_GDNLine_P
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


