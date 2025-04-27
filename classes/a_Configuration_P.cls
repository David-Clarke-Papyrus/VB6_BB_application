VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Configuration_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Save(ByVal buffer As String, pStatus As String) As String
    On Error GoTo errHandler

Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ConfigProps
Dim udtData As ConfigData
Dim objPB As PropertyBag
Dim objPersist As a_Company_P
Dim objPersistCur As a_Currency_P
Dim objPersistStores As a_store_P
Dim objPersistStaff As a_Staff_P
Dim objPersistDCs As a_DocumentControl_P
Dim objPersistRRs As a_RoundingRule_P
Dim objPersist5 As c_BICCodes_P
Dim oBudgets As ADODB.Recordset
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim lngIndex As Long
Dim lngCount As Long
Dim strStatus As String

  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    udtData.buffer = .ReadProperty("State")
  End With
  LSet udtProps = udtData
  
  strSQL = "SELECT * FROM tConfiguration"
  Set rs = New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
  rs.open strSQL, oPC.COShort, adOpenKeyset, adLockOptimistic
  If udtProps.IsNew Then rs.AddNew

  With rs
  On Error Resume Next
        .fields("CF_Prefix") = Trim$(udtProps.TransactionPrefix)
'        .Fields("CF_WS_Start") = CVDate(udtProps.WSStart)
'        .Fields("CF_MS_Start") = CVDate(udtProps.MSStart)
        .fields("CF_TFRDiscount") = udtProps.TFRDiscount
        .fields("CF_TFRDiscountAdj") = FNN(udtProps.TFRDiscountAdj)
        .fields("CF_VATRate") = udtProps.VATRate
        .fields("CF_LocalCountryID") = udtProps.LocalCountryID
        .fields("CF_DefaultCurrID") = udtProps.DefaultCurrencyID
        .fields("CF_LocalCurrID") = udtProps.LocalCurrencyID
        .fields("CF_DefaultCompanyID") = udtProps.DefaultCOMPID
        .fields("CF_DefaultStoreID") = udtProps.DefaultStoreID
        .fields("CF_BillToStoreID") = udtProps.BillToStoreID
        .fields("CF_ShowProdsWithInstancesOnly") = FNB(udtProps.ShowProdsWithInstancesOnly)
        .fields("CF_OrderText") = FNS((udtProps.OrderText))
        .fields("CF_SalesOrderText") = FNS((udtProps.SalesOrderText))
        .fields("CF_QuotationText") = FNS((udtProps.QuotationText))
        .fields("CF_InvoiceText") = FNS((udtProps.InvoiceText))
        .fields("CF_StatementText") = FNS((udtProps.StatementText))
        .fields("CF_LastStkTkeDate") = FND(udtProps.LastStockTakeDate)
        .fields("CF_AllowCopyInfo") = FNB(udtProps.AllowCopyInfo)
        .fields("CF_AntiquarianYN") = FNB(udtProps.AntiquarianYN)
        .fields("CF_CasualCustomersYN") = FNB(udtProps.CasualCustomersYN)
        .fields("CF_NonBookYN") = FNB(udtProps.NonBookYN)
        .fields("CF_EnforceSections") = FNB(udtProps.EnforceSections)
        .fields("CF_SignTransactions") = FNB(udtProps.SignTransactions)
        .fields("CF_CaptureDecimalYN") = FNB(udtProps.CaptureDecimal)
        .fields("CF_OfferSignature") = FNS(udtProps.OfferSignature)
        .fields("CF_LookupSeq") = FNS(udtProps.LookupSeq)
'        .Fields("CF_DaysInWeek") = FNN(udtProps.DaysInWeek)
'        .Fields("CF_UpdateWindowStart") = FND(udtProps.UpdateWindowStart)
'        .Fields("CF_UpdateWindowEnd") = FND(udtProps.UpdateWindowEnd)
'        .Fields("CF_OpsetsAuto") = FNN(udtProps.OpSetsAuto)
        .fields("CF_LastWantsExportDate") = FND(udtProps.LastWantsExportDate)
        .fields("CF_INVATRegion") = FNB(udtProps.IsVATRegion)
        .fields("CF_DiscountVAT") = FNB(udtProps.DiscountVAT)
        .fields("CF_EnforceCOLRef") = FNB(udtProps.EnforceCOLRef)
        .fields("CF_SupportsWORD") = FNB(udtProps.SupportsWORD)
        .fields("CF_COTypesSupported") = FNN(udtProps.COTypesSupported)
        .fields("CF_CSCustomerID") = FNN(udtProps.CSCustomerID)
        .fields("CF_PrintingSettings") = FNS(udtProps.PrintingSettings)
        .fields("CF_UNallocatedPT") = FNN(udtProps.DefaultPT)
        .fields("CF_UnallocatedCT") = FNN(udtProps.DefaultCT)
        .fields("CF_DefaultSection") = FNN(udtProps.DefaultSection)
        .fields("CF_COLAllocationStyle") = FNS(udtProps.COLAllocationStyle)
        .fields("CF_ReorderperCOL") = FNB(udtProps.ReorderPerCOL)
        .fields("CF_AggregatePOs") = FNB(udtProps.AggregatePOs)
        .fields("CF_SupportsLoyaltyClub") = FNB(udtProps.SupportsLoyaltyClub)
        .fields("CF_GFXNumber") = FNS(udtProps.GFXNumber)
        .fields("CF_CreditorsContraAccount") = FNS(udtProps.IE_CreditorsContraAccount)
        .fields("CF_DebtorsContraAccount") = FNS(udtProps.IE_DebtorsContraAccount)
        .fields("CF_GLReference") = FNS(udtProps.IE_GLReference)
        .fields("CF_AccountingApplicationName") = FNS(udtProps.IE_AccountingApplicationName)
        .fields("CF_MinMU") = FNN(udtProps.MinMU)
        .fields("CF_LoyStartNumber") = FNN(udtProps.LoyStartNumber)
        .fields("CF_LoyEndNumber") = FNN(udtProps.LoyEndNumber)
        .fields("CF_EmailPOMsg") = FNS(udtProps.EmailPOMsg)
        .fields("CF_EmailInvMsg") = FNS(udtProps.EmailInvMsg)
        .fields("CF_EmailQuoteMsg") = FNS(udtProps.EmailQuoteMsg)
        .fields("CF_EmailAppMsg") = FNS(udtProps.EmailAPPMsg)
        
    On Error GoTo errHandler
        .Update
        udtProps.IsNew = False
        udtProps.IsDirty = False
    End With
    rs.Close
  Set rs = Nothing
  
  Set objPBOut = New PropertyBag
  
        LSet udtData = udtProps
        objPBOut.WriteProperty "State", udtData.buffer
        
        Set objPersist = New a_Company_P
        objPBOut.WriteProperty "Companies", objPersist.Save(objPB.ReadProperty("Companies"))

        Set objPersistCur = New a_Currency_P
        objPBOut.WriteProperty "Currencies", objPersistCur.Save(objPB.ReadProperty("Currencies"))
        
        Set objPersistStores = New a_store_P
        objPBOut.WriteProperty "Stores", objPersistStores.Save(objPB.ReadProperty("Stores"))

        Set objPersistStaff = New a_Staff_P
        objPBOut.WriteProperty "Staff", objPersistStaff.Save(objPB.ReadProperty("Staff"), strStatus)
        If strStatus > "" Then pStatus = strStatus
        
        Set objPersistRRs = New a_RoundingRule_P
        objPBOut.WriteProperty "RRs", objPersistRRs.Save(objPB.ReadProperty("RRs"))
        
        Set objPersistDCs = New a_DocumentControl_P
        objPBOut.WriteProperty "DCs", objPersistDCs.Save(objPB.ReadProperty("DCs"))
        
        Set objPersist5 = New c_BICCodes_P
        objPBOut.WriteProperty "BICs", objPersist5.Fetch  'We never change these from the Gui but we might reload them after importing from the CD

        Set objPB = Nothing
        Set objPersist = Nothing
        
        Save = objPBOut.Contents
        Set objPBOut = Nothing
        Set objPersistCur = Nothing
        Set objPersistStores = Nothing
        Set objPersistStaff = Nothing
        Set objPersistRRs = Nothing
        Set objPersist5 = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
        
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration_P.Save(buffer,pStatus)", Array(buffer, pStatus)
End Function
Public Function Fetch() As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim udtProps As ConfigProps
Dim udtData As ConfigData
Dim objPB As PropertyBag
Dim lngCount As Long
Dim objPersist As a_Company_P
Dim objPersist2 As a_Currency_P
Dim objPersist3 As a_store_P
Dim objPersist4 As a_RoundingRule_P
Dim objPersist5 As c_BICCodes_P
Dim objPersist6 As a_Staff_P
Dim objPersist7 As a_DocumentControl_P
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    rs.open "SELECT * FROM tConfiguration", oPC.COShort, adOpenStatic, adLockReadOnly
    Set objPB = New PropertyBag
    If Not rs.eof Then
        With rs
        On Error Resume Next
        udtProps.TransactionPrefix = FNS(.fields("CF_Prefix"))
'        udtProps.WSStart = FND(.Fields("CF_WS_Start"))
'        udtProps.MSStart = FND(.Fields("CF_MS_Start"))
        udtProps.TFRDiscount = FNN(.fields("CF_TFRDiscount"))
        udtProps.TFRDiscountAdj = FNN(.fields("CF_TFRDiscountAdj"))
        udtProps.ConfirmBeforePrint = FNB(.fields("CF_PromptBeforePrint"))
        udtProps.AllowCopyInfo = FNB(.fields("CF_AllowCopyInfo"))
        udtProps.AntiquarianYN = FNB(.fields("CF_AntiquarianYN"))
        udtProps.DiscountVATYN = FNB(.fields("CF_DiscountVAT"))
        udtProps.CasualCustomersYN = FNB(.fields("CF_CasualCustomersYN"))
        udtProps.NonBookYN = FNB(.fields("CF_NonBookYN"))
        udtProps.EnforceSections = FNB(.fields("CF_EnforceSections"))
        udtProps.SignTransactions = FNB(.fields("CF_SignTransactions"))
        udtProps.CaptureDecimal = FNB(.fields("CF_CaptureDecimalYN"))
        udtProps.ShowProdsWithInstancesOnly = FNB(.fields("CF_ShowProdsWithInstancesOnly"))
        udtProps.LocalCountryID = FNN(.fields("CF_LocalCountryID"))
        udtProps.DefaultCurrencyID = FNN(.fields("CF_DefaultCurrID"))
        udtProps.LocalCurrencyID = FNN(.fields("CF_LocalCurrID"))
        udtProps.DefaultCOMPID = FNN(.fields("CF_DefaultCompanyID"))
        udtProps.DefaultStoreID = FNN(.fields("CF_DefaultStoreID"))
        udtProps.BillToStoreID = FNN(.fields("CF_BillToStoreID"))
        udtProps.DefaultPT = FNN(.fields("CF_UnallocatedPT"))
        udtProps.DefaultCT = FNN(.fields("CF_UnallocatedCT"))
        udtProps.DefaultSection = FNN(.fields("CF_DefaultSection"))
        udtProps.OrderText = FNS(.fields("CF_OrderText"))
        udtProps.QuotationText = FNS(.fields("CF_QuotationText"))
        udtProps.SalesOrderText = FNS(.fields("CF_SalesOrderText"))
        udtProps.InvoiceText = FNS(.fields("CF_InvoiceText"))
        udtProps.StatementText = FNS(.fields("CF_StatementText"))
        udtProps.LastStockTakeDate = FND(.fields("CF_LastStkTkeDate"))
        udtProps.OfferSignature = FNS(.fields("CF_OfferSignature"))
        udtProps.LookupSeq = FNS(.fields("CF_LookupSEq"))
        udtProps.LastWantsExportDate = FND(.fields("CF_LastwantsExportDate"))
        udtProps.VATRate = FNDBL(.fields("CF_VATRate"))
        udtProps.IsVATRegion = FNB(.fields("CF_InVATRegion"))
        udtProps.ReorderPerCOL = FNB(.fields("CF_ReorderPerCOL"))
        udtProps.EnforceCOLRef = FNB(.fields("CF_EnforceCOLRef"))
        udtProps.AggregatePOs = FNB(.fields("CF_AggregatePOs"))
        udtProps.SupportsWORD = FNB(.fields("CF_SupportsWORD"))
        udtProps.DiscountVAT = FNB(.fields("CF_DiscountVAT"))
        udtProps.COTypesSupported = FNN(.fields("CF_COTypesSupported"))
        udtProps.CSCustomerID = FNN(.fields("CF_CSCustomerID"))
        udtProps.PrintingSettings = FNS(.fields("CF_PrintingSettings"))
        udtProps.COLAllocationStyle = FNS(.fields("CF_COLAllocationStyle"))
        udtProps.LoyaltyClubType = FNN(.fields("CF_LoyaltyClubType"))
        udtProps.SupportsLoyaltyClub = FNB(.fields("CF_SupportsLOyaltyClub"))
        udtProps.CustomerTypeDictID = FNN(.fields("CF_CustomerTypeDictID"))
        udtProps.InterestGroupsDictID = FNN(.fields("CF_CustomerIGDictID"))
        udtProps.InterestGroupsDictID = FNN(.fields("CF_CustomerIGDictID"))
        udtProps.SectionTypeDictID = FNN(.fields("CF_SectionDictID"))
        udtProps.BusinessCustomerTypeID = FNN(.fields("CF_BusinessCustomerTypeID"))
        udtProps.BookClubCustomerTypeID = FNN(.fields("CF_BC_ID"))
        udtProps.PrivateCustomerTypeID = FNN(.fields("CF_PrivateCustomerTypeID"))
        udtProps.GFXNumber = FNS(.fields("CF_GFXNumber"))
        udtProps.IE_CreditorsContraAccount = FNS(.fields("CF_CreditorsContraAccount"))
        udtProps.IE_DebtorsContraAccount = FNS(.fields("CF_DebtorsContraAccount"))
        udtProps.IE_GLReference = FNS(.fields("CF_GLReference"))
        udtProps.IE_AccountingApplicationName = FNS(.fields("CF_AccountingApplicationName"))
        udtProps.LastUpdateDate = FND(.fields("CF_LastUpdateDate"))
        udtProps.MinMU = FNN(.fields("CF_MinMU"))
        udtProps.LoyStartNumber = FNN(.fields("CF_LoyStartNumber"))
        udtProps.LoyEndNumber = FNN(.fields("CF_LoyEndNumber"))
        udtProps.EmailPOMsg = FNS(.fields("CF_EmailPOMsg"))
        udtProps.EmailInvMsg = FNS(.fields("CF_EmailInvMsg"))
        udtProps.EmailQuoteMsg = FNS(.fields("CF_EmailQuoteMsg"))
        udtProps.EmailAPPMsg = FNS(.fields("CF_EmailAppMsg"))
        udtProps.StatementAsAt = FND(.fields("CF_StatementAsAt"))
        
        
    On Error GoTo errHandler
        LSet udtData = udtProps
        End With
        rs.Close
        Set rs = Nothing
        
        Set objPersist = New a_Company_P
        Set objPB = New PropertyBag
        With objPB
          .WriteProperty "State", udtData.buffer
          .WriteProperty "Companies", objPersist.Fetch
          Fetch = .Contents
        End With
        Set objPersist = Nothing
        
        Set objPersist2 = New a_Currency_P
        With objPB
            .WriteProperty "Currencies", objPersist2.Fetch
            Fetch = .Contents
        End With
        Set objPersist3 = New a_store_P
        With objPB
            .WriteProperty "Stores", objPersist3.Fetch
            Fetch = .Contents
        End With
        Set objPersist4 = New a_RoundingRule_P
        With objPB
            .WriteProperty "RRs", objPersist4.Fetch
            Fetch = .Contents
        End With
        Set objPersist5 = New c_BICCodes_P
        With objPB
            .WriteProperty "BICs", objPersist5.Fetch
            Fetch = .Contents
        End With
        Set objPersist6 = New a_Staff_P
        With objPB
            .WriteProperty "Staff", objPersist6.Fetch
            Fetch = .Contents
        End With
        Set objPersist7 = New a_DocumentControl_P
        With objPB
            .WriteProperty "DCs", objPersist7.Fetch
            Fetch = .Contents
        End With
        Set objPB = Nothing
        Set objPersist2 = Nothing
        Set objPersist3 = Nothing
        Set objPersist4 = Nothing
        Set objPersist5 = Nothing
        Set objPersist6 = Nothing
        Set objPersist7 = Nothing

    Else
        rs.MoveNext
  End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
  Exit Function
        
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration_P.Fetch"
End Function

Private Sub DeleteConfiguration()
    On Error GoTo errHandler
Dim strSQL As String

  strSQL = "DELETE FROM tblConfiguration WHERE ID=1"
  oPC.COShort.execute strSQL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration_P.DeleteConfiguration"
End Sub

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Configuration_P.DeleteObject(ID)", ID
End Sub

