VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_QU_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function Fetch(TRID As Long) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
  
Dim strSQL As String
Dim udtProps As QUProps
Dim udtData As QUData
Dim objPersist As a_QUL_P
Dim objPB As PropertyBag
Dim oUtil As z_SQL
Dim OpenResult As Integer

    Set rs = New ADODB.Recordset
    Set oUtil = New z_SQL
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngResult = oUtil.RunGetRecordset("SELECT * FROM Quote_Fetch WHERE TR_ID = " & TRID, enText, Array(), "", rs)
    
    
    If Not rs.eof Then
        With udtProps
            .QuoteID = rs!TR_ID
            .TPID = FNN(rs!TR_TP_ID)
            .BillToAddressID = FNN(rs.fields("AddressDOCID"))
            .DelToAddressID = FNN(rs.fields("AddressGOODSID"))
            .Ref = FNS(rs.fields("QU_ORDERNUM"))
            .OrderDate = FND(rs.fields("QU_ORDERDate"))
            .ForAttn = FNS(rs.fields("QU_ForAttn"))
           ' .CurrencyID_Foreign = FNN(rs!QU_Curr_ID)
            .Status = FNN(rs.fields("TR_Status"))
           ' .VATRate = FNDBL(rs!QU_VATRate)
           ' .DiscountRate = FNDBL(rs!QU_DiscountPercent)
           ' .TotalDiscount = FNN(rs!QU_DiscountAmount)
           ' .TotalVAT = FNN(rs!QU_VATAmount)
           ' .TotalPayable = FNN(rs!QU_PayableAmount)
            .COMPID = FNN(rs!TR_COMP_ID)
           ' .ShowVAT = FNB(rs!QU_ShowVAT)
           ' .CurrencyFormat = FNS(rs!CURR_Format)
            .BusPhone = FNS(rs!ADD_BusPhone)
            .TPACCNum = FNS(rs!TP_ACno)
            .Log = Trim(IIf(IsNull(rs!TR_Log), "", rs!TR_Log))
            .TPFax = FNS(rs!ADD_Fax)
            .TPPhone = FNS(rs!TP_Phone)
            .TPNAME = FNS(rs!TP_Name)
            .Memo = FNS(rs!TR_NOTE)
            .DOCCode = FNS(rs!TR_CODE)
            '.Proforma = FNB(rs!QU_Proforma)
            .DOCDate = FND(rs!TR_DATE)
            .CaptureDate = FND(rs!TR_CaptureDate)
            .StaffID = FNN(rs!TR_StaffID)
            If .StaffID > 0 Then
                .StaffName = oPC.Configuration.Staff.FindStaffByID(.StaffID).Shortname
                .Signature = oPC.Configuration.Staff.FindStaffByID(.StaffID).Signature
                .StaffEmail = oPC.Configuration.Staff.FindStaffByID(.StaffID).EMail
            End If
            
           ' .SalesRepID = FNN(rs!QU_REPID)
           ' .SalesRepName = FNS(rs!SalesRepName)
           ' .CustPaid = FNB(rs!QU_InvoicePaid)
           ' .CommPaid = FNB(rs!QU_CommissionPaid)
           ' .VATable = FNB(rs!QU_Vatable)
           ' .ShowVAT = FNB(rs!QU_ShowVAT)
            .IsNew = False
            .IsDirty = False
            .IsDeleted = False
            rs.Close
            Set rs = Nothing
            Set oUtil = Nothing
            LSet udtData = udtProps
              Set objPersist = New a_QUL_P
            Set objPB = New PropertyBag
            With objPB
              .WriteProperty "State", udtData.buffer
              .WriteProperty "QULS", objPersist.Fetch(udtProps.QuoteID)
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
    ErrorIn "a_QUL_P.Fetch(TRID)", TRID
    Exit Function
    Resume Next
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim bUpdateStock As Boolean
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim objCode As New z_Code
Dim udtProps As QUProps
Dim udtData As QUData
Dim objPersist As a_QUL_P
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
    strSQL = "SELECT * FROM tTR WHERE TR_ID =" & udtProps.QuoteID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    With udtProps
        If .IsNew Then
            rs.AddNew
        End If
        rs!TR_Type = 16
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
        rs!TR_COMP_ID = oPC.Configuration.DefaultCOMPID
        rs.Update
        If .IsNew Then
           rs.MoveLast
           .QuoteID = rs!TR_ID
        End If
        rs.Close
    End With

    strSQL = "SELECT * FROM tQU WHERE QU_ID =" & udtProps.QuoteID
    rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
    
    With udtProps
        If udtProps.IsNew Then rs.AddNew
        rs!QU_ID = FNN(.QuoteID)
        rs.fields("QU_ORDERNUM") = FNS(.Ref)
        rs.fields("QU_ORDERDate") = FND(.OrderDate)
        rs.fields("QU_ForAttn") = FNS(.ForAttn)

      '  rs!QU_Curr_ID = FNN(.CurrencyID_Foreign)
      '  rs!QU_Proforma = FNB(.Proforma)
      '  rs!QU_REPID = FNN(.SalesRepID)
      '  rs!QU_CUrrRate = FNDBL(.CurrencyFactor)  '   Conversion rate
      '  rs!QU_DiscountPercent = FNDBL(.DiscountRate)  '   Invoice discount (general) rate
      '  rs!QU_VATRate = FNDBL(.VATRate)  '
      '  rs!QU_Vatable = FNB(.VATable)
      '  rs!QU_DiscountAmount = FNDBL(.TotalDiscount) / oPC.Configuration.DefaultCurrency.Divisor
      '  rs!QU_VATAmount = FNDBL(.TotalVAT) / oPC.Configuration.DefaultCurrency.Divisor
      '  rs!QU_PayableAmount = FNDBL(.TotalPayable) / oPC.Configuration.DefaultCurrency.Divisor
      '  rs!QU_ShowVAT = FNB(.ShowVAT)
        rs.Update
        .IsNew = False
        .IsDirty = False
    End With
    rs.Close
    Set rs = Nothing
    
    Set objPBOut = New PropertyBag
    LSet udtData = udtProps
    Set objPersist = New a_QUL_P
      With objPBOut
        .WriteProperty "State", udtData.buffer
        .WriteProperty "QULS", objPersist.Save(objPB.ReadProperty("QULS"), udtProps.QuoteID)
      End With
      Set objPB = Nothing
      Set objPersist = Nothing
    
      Save = objPBOut.Contents
      Set objPBOut = Nothing
      
        
   ' oPC.COShort.execute "Execute dbo.AllocateSalesCommission " & udtProps.QuoteID & "," & FNN(udtProps.SalesRepID)
  
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
    ErrorIn "a_QUL_P.Save(buffer)", buffer, , rs
End Function


Private Sub DeleteInvoice(ID As Long)
    On Error GoTo errHandler
'    oPC.COShort.execute "DELETE FROM tPI WHERE PI_ID=" & ID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL_P.DeleteInvoice(ID)", ID
End Sub

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
  Dim objPersist As a_QUL_P
  
    oPC.COShort.BeginTrans
    oPC.COShort.execute "DELETE FROM tQU WHERE QU_ID=" & ID
  
    Set objPersist = New a_QUL_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing
    oPC.COShort.CommitTrans
    Exit Sub
  
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_QUL_P.DeleteObject(ID)", ID
End Sub


