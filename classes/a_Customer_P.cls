VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Customer_P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements ObjectControl

'Private mobjContext As ObjectContext

Public Function Fetch(Optional CustomerID As Long, Optional pACNO As String, Optional pDefaultPhone As String, Optional pSAN As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim cmd As New ADODB.Command
Dim strSQL As String
Dim strWHERE As String
Dim udtProps As CustomerProps
Dim udtData As CustomerData
Dim objPersist As a_Address_P
Dim objPersist2 As a_IG_P
Dim objPersist3 As a_IG_P
Dim objPB As PropertyBag
Dim OpenResult As Integer
  
    If CustomerID > 0 Then
        strWHERE = " a.TP_ID = " & CustomerID
    ElseIf Len(pACNO) > 0 Then
        strWHERE = " a.TP_ACNO = '" & pACNO & "'"
    ElseIf Len(pSAN) > 0 Then
        strWHERE = " a.TP_SAN = '" & pSAN & "'"
    ElseIf Len(pDefaultPhone) > 0 Then
        strWHERE = " a.TP_Phone = " & pDefaultPhone
    End If
    strSQL = "SELECT DISTINCT a.*, tADD.*, tStore.* ,SM_Name as REPNAME,b.TP_NAME ParentCustomerName FROM tTP a " & _
            "LEFT  JOIN tSTORE ON a.TP_LoyaltyHomeStoreID = Store_ID " & _
            "Left  JOIN tADD ON a.TP_ID=tADD.ADD_TP_ID " & _
            "LEFT JOIN tStaffMember on a.TP_REP_ID = SM_ID " & _
            "LEFT JOIN tTP b ON a.TP_PARENTTPID = b.TP_ID " & _
            "WHERE " & strWHERE

    Set rs = New ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.open strSQL, oPC.COShort, adOpenDynamic, adLockOptimistic
    If Not rs.eof Then
        With rs
          udtProps.CustID = !TP_ID
          udtProps.CustomerTypeID = FNN(!TP_CT_ID)
          udtProps.ParentCustomerID = FNN(!TP_ParentTPID)
          udtProps.ParentCustomerName = FNS(!ParentCustomerName)
          udtProps.Initials = FNS(!TP_Initials)
          udtProps.MOBILE = FNS(!TP_Cell)
          udtProps.Title = FNS(!TP_Title)
          udtProps.DefaultAddressID = FNN(!TP_DefaultAddressID)
          udtProps.Name = FNS(!TP_Name)
          udtProps.Note = FNS(!TP_Note)
          udtProps.Phone = FNS(!TP_Phone)
          udtProps.StoreName = FNS(!STORE_Name)
          udtProps.StoreID = FNN(!STORE_ID)
          udtProps.SearchPhone = FNS(!TP_SearchPhone)
          udtProps.VATable = FNB(!TP_Vatable)
          udtProps.VatNumber = FNS(!TP_VATNumber)
          udtProps.GetsCatalogue = FNB(!TP_GetsCatalogue)
          udtProps.CustNotifyBookLaunch = FNB(!TP_CustNotifyBookLaunch)
          udtProps.CustNotifyBookPromotion = FNB(!TP_CustNotifyBookPromotion)
          udtProps.CustNotifyBookSale = FNB(!TP_CustNotifyBookSale)
          udtProps.AcNo = FNS(!TP_ACno)
          udtProps.SAN = FNS(!TP_SAN)
          udtProps.DefaultDiscount = FNDBL(!TP_DefaultDiscount)
          udtProps.CreditLimit = FNN(!TP_CreditLimit)
          udtProps.Terms = FNN(!TP_Terms)
          udtProps.BalCur = FNDBL(!TP_BALANCE_CUR)
          udtProps.Bal30 = FNDBL(!TP_BALANCE_30)
          udtProps.Bal60 = FNDBL(!TP_BALANCE_60)
          udtProps.Bal90 = FNDBL(!TP_BALANCE_90)
          udtProps.Bal120 = FNDBL(!TP_BALANCE_120)
          udtProps.Bal120Plus = FNDBL(!TP_BALANCE_120Plus)
          udtProps.Balance = FNDBL(!TP_Balance)
          udtProps.PaymentStyle = FNS(!TP_PAYMENTSTYLE)
          udtProps.Blocked = FNB(!TP_Blocked)
          udtProps.oLineToInvoice = IIf(IsNull(!TP_olineNotetoInvoice) = True, True, FNB(!TP_olineNotetoInvoice))
          udtProps.UseQuotedPrice = FNB(!TP_UseQuotedPrice)
          udtProps.OurACnoWithClient = FNS(!TP_OurACnumWithClient)
          udtProps.AccAcno = FNS(!TP_ACCACNO)
          udtProps.ShowVAT = FNB(!TP_ShowVAT)
          udtProps.CanBeDeleted = FNB(!TP_CanBeDeletedYN)
          udtProps.OneLinePerInvoice = FNB(!TP_OneLinePerInvoice)
          udtProps.Role = FNN(!TP_Role)
          udtProps.SalesRepID = FNN(!TP_REP_ID)
          udtProps.Repname = FNS(!Repname)
          If FNS(udtProps.Repname) = "" Then udtProps.Repname = "<NONE>"
          udtProps.SalesOrderTemplateName = FNS(!TP_SalesOrderTemplateName)
          udtProps.QuotationTemplateName = FNS(!TP_QuotationTemplateName)
          udtProps.ApproTemplateName = FNS(!TP_ApproTemplateName)
          udtProps.ApproReturnTemplateName = FNS(!TP_ApproReturnTemplateName)
          udtProps.InvoiceTemplateName = FNS(!TP_InvoiceTemplateName)
       ''   udtProps.GDNTemplateName = FNS(!TP_GDNTemplateName)
          udtProps.CreditNoteTemplateName = FNS(!TP_CreditNoteTemplateName)
          udtProps.ContactPerson = FNS(!TP_ContactPerson)
          udtProps.ContactpersonPhone = FNS(!TP_ContactpersonPhone)
          udtProps.DispatchMethod = FNS(!TP_DispatchMethod)
          udtProps.DateLastModified = FND(!TP_DateLastModified)
          udtProps.DateRecordAdded = FND(!TP_DateRecordAdded)
          udtProps.OpenItemOrBalBF = FNS(!TP_OpenItemOrBalBF)
          udtProps.GenerateSeparateInvoicesForSeparateOrders = FNB(!TP_GenerateSeparateInvoicesForSeparateOrders)
          udtProps.CompleteOrder = FNB(!TP_CompleteOrder)
          .Close
        End With
        Set rs = Nothing
        LSet udtData = udtProps
        
        Set objPersist = New a_Address_P
        Set objPersist2 = New a_IG_P
        Set objPersist3 = New a_IG_P
        Set objPB = New PropertyBag
        With objPB
          .WriteProperty "State", udtData.buffer
          .WriteProperty "Addresses", objPersist.Fetch(udtProps.CustID)
          .WriteProperty "IGs", objPersist2.Fetch(udtProps.CustID, oPC.Configuration.InterestGroupsDictID)
          .WriteProperty "CCs", objPersist3.Fetch(udtProps.CustID, oPC.Configuration.CustomerTypeDictID)
          Fetch = .Contents
        End With
        Set objPB = Nothing
        Set objPersist = Nothing
        Set objPersist2 = Nothing
        Set objPersist3 = Nothing
  Else
    ' force an or
        Fetch = ""
  End If
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer_P.Fetch(CustomerID,pACNO,pDefaultPhone)", Array(CustomerID, pACNO, _
         pDefaultPhone)
End Function

Public Function Save(ByVal buffer As String) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim rsRole As ADODB.Recordset
Dim udtProps As CustomerProps
Dim udtData As CustomerData
Dim objPersist As a_Address_P
Dim objPersist2 As a_IG_P
Dim objPersist3 As a_IG_P
Dim objPB As PropertyBag
Dim objPBOut As PropertyBag
Dim arBuffer() As Byte
Dim tmpBuffer As String
Dim strSema As String
Dim OpenResult As Integer
Dim cmd As New ADODB.Command
Dim strSQL As String
    
 '   oPC.COShort.BeginTrans
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set objPB = New PropertyBag
    arBuffer = buffer
    With objPB
      .Contents = arBuffer
      udtData.buffer = .ReadProperty("State")
    End With
    LSet udtProps = udtData
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    With rs
        .open "SELECT * FROM tTP WHERE TP_ID = " & udtProps.CustID, oPC.COShort, adOpenDynamic, adLockOptimistic
        If udtProps.IsNew Then rs.AddNew
        !TP_Initials = FNS(udtProps.Initials)
        !TP_Title = FNS(udtProps.Title)
        !TP_Name = FNS(udtProps.Name)
        !TP_Note = FNS(udtProps.Note)
        !TP_OnMailList = False
        !TP_Cell = FNS(udtProps.MOBILE)
        !TP_Role = 3
        !TP_LoyaltyHomeStoreID = FNN(udtProps.StoreID)
        !TP_CanBeDeletedYN = FNB(udtProps.CanBeDeleted)
        !TP_Vatable = FNB(udtProps.VATable)
        !TP_VATNumber = FNS(udtProps.VatNumber)
        !TP_GetsCatalogue = FNB(udtProps.GetsCatalogue)
        !TP_DefaultDeliveryDays = 30
        !TP_DefaultDiscount = FNDBL(udtProps.DefaultDiscount)
        !TP_CT_ID = udtProps.CustomerTypeID
'        If Right(FNS(udtProps.AcNo), 1) = "." Then
'            udtProps.AcNo = GetNextAvailableTPAcno(Left(FNS(udtProps.AcNo), Len(FNS(udtProps.AcNo)) - 1))
'        End If
        !TP_ACno = FNS(udtProps.AcNo)
        !TP_SAN = FNS(udtProps.SAN)
        !TP_DefaultAddressID = udtProps.DefaultAddressID
        !TP_Phone = FNS(udtProps.Phone)
        !TP_SearchPhone = (FNS(udtProps.SearchPhone))
        !TP_DateLastModified = ReverseDateTime(Now())
        !TP_CustNotifyBookSale = FNB(udtProps.CustNotifyBookSale)
        !TP_CustNotifyBookPromotion = FNB(udtProps.CustNotifyBookPromotion)
        !TP_CustNotifyBookLaunch = FNB(udtProps.CustNotifyBookLaunch)
        !TP_DispatchMethod = FNS(udtProps.DispatchMethod)
        !TP_REP_ID = FNN(udtProps.SalesRepID)
        !TP_CreditLimit = FNN(udtProps.CreditLimit)
        !TP_Terms = FNN(udtProps.Terms)
'        !TP_TermsType = FNS(udtProps.TermsType)

        !TP_PAYMENTSTYLE = FNS(udtProps.PaymentStyle)
        !TP_Blocked = FNB(udtProps.Blocked)
        !TP_olineNotetoInvoice = FNB(udtProps.oLineToInvoice)
        !TP_UseQuotedPrice = FNB(udtProps.UseQuotedPrice)
        !TP_OurACnumWithClient = FNS(udtProps.OurACnoWithClient)
        !TP_ACCACNO = FNS(udtProps.AccAcno)
        !TP_ShowVAT = FNB(udtProps.ShowVAT)
        !TP_OneLinePerInvoice = FNB(udtProps.OneLinePerInvoice)
        !TP_SalesOrderTemplateName = FNS(udtProps.SalesOrderTemplateName)
        !TP_QuotationTemplateName = FNS(udtProps.QuotationTemplateName)
        !TP_ApproTemplateName = FNS(udtProps.ApproTemplateName)
        !TP_ApproReturnTemplateName = FNS(udtProps.ApproReturnTemplateName)
        !TP_InvoiceTemplateName = FNS(udtProps.InvoiceTemplateName)
        !TP_GDNTemplateName = FNS(udtProps.GDNTemplateName)
        !TP_CreditNoteTemplateName = FNS(udtProps.CreditNoteTemplateName)
        !TP_ParentTPID = FNN(udtProps.ParentCustomerID)
        !TP_ContactPerson = FNS(udtProps.ContactPerson)
        !TP_ContactpersonPhone = FNS(udtProps.ContactpersonPhone)
        !TP_OpenItemOrBalBF = FNS(udtProps.OpenItemOrBalBF)
        !TP_GenerateSeparateInvoicesForSeparateOrders = FNB(udtProps.GenerateSeparateInvoicesForSeparateOrders)
        !TP_CompleteOrder = FNB(udtProps.CompleteOrder)
   '     !TP_SettlementDiscount = FNDBL(udtProps.SettlementDiscount)
   '     !TP_SettlementTerms = FNN(udtProps.SettlementTerms)
   '     !TP_SettlementTermsType = FNS(udtProps.SettlementTermsType)
   '     !TP_GenerateStatement = FNB(udtProps.GenerateStatement)
        .Update
        If udtProps.IsNew Then
            rs.MoveLast '            rs.Bookmark = rs.Bookmark
            udtProps.CustID = !TP_ID
        End If
        
    Set objPBOut = New PropertyBag
  
    LSet udtData = udtProps
  
    Set objPersist = New a_Address_P
    Set objPersist2 = New a_IG_P
    Set objPersist3 = New a_IG_P
    With objPBOut
        .WriteProperty "State", udtData.buffer
        tmpBuffer = objPersist.Save(objPB.ReadProperty("Addresses"), udtProps.CustID)
        If tmpBuffer > "" Then
            .WriteProperty "Addresses", tmpBuffer
        Else
            strSema = "Duplicate value in Address"
            GoTo errHandler
        End If
        tmpBuffer = objPersist2.Save(objPB.ReadProperty("IGs"), udtProps.CustID)
        If tmpBuffer > "" Then
            .WriteProperty "IGs", tmpBuffer
        End If
        tmpBuffer = objPersist3.Save(objPB.ReadProperty("CCs"), udtProps.CustID)
        If tmpBuffer > "" Then
            .WriteProperty "CCs", tmpBuffer
        End If
    End With
    Set objPB = Nothing
    Set objPersist = Nothing
    Set objPersist2 = Nothing
    Set objPersist3 = Nothing

    Save = objPBOut.Contents
    Set objPBOut = Nothing
    rs.Close
    Set rs = Nothing
        
   End With
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
   
EXITH:
    Exit Function
errHandler:
    ErrPreserve
    If strSema = "Duplicate value in Address" Then
        Set rs = Nothing
        Save = ""
        Exit Function
    ElseIf Error = "DUPLICATE" Then    'There is an attempted duplication
        Set rs = Nothing
        Save = "DUPLICATE"
        Exit Function
    End If
    RlsObjs rs
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer_P.Save(buffer)", buffer, EA_ROLLBACK + EA_RERAISE + EA_ADVANCED, oPC.COShort
    'Resume
End Function

Public Sub DeleteObject(ByVal ID As Long)
    On Error GoTo errHandler
Dim objPersist As a_Address_P

    Set objPersist = New a_Address_P
    objPersist.DeleteObject ID
    Set objPersist = Nothing

    oPC.COShort.execute "DELETE FROM tTP WHERE TP_ID= " & ID

  Exit Sub

'H:
'    MsgBox "A-Customer_P: DeleteObject: " & or
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Customer_P.DeleteObject(ID)", ID
End Sub



'Public Function AddIG(pTPID As Long, pIGID As Long)
'    On Error GoTo ErrHandler
'Dim oUtil As z_UTIL_P
'    Set oUtil = New z_UTIL_P
'    oUtil.RunSQL "INSERT INTO tTP_IG (TPIG_TP_ID,TPIG_IG_ID) VALUES (" & pTPID & "," & pIGID & ")"
'    Set oUtil = Nothing
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer_P.AddIG(pTPID,pIGID)", Array(pTPID, pIGID)
'End Function
'Public Function RemoveIG(pTPID As Long, pIGID As Long)
'    On Error GoTo ErrHandler
'Dim oUtil As z_UTIL_P
'    Set oUtil = New z_UTIL_P
'    oUtil.RunSQL "DELETE FROM tTP_IG WHERE TPIG_TP_ID = " & pTPID & " AND IG_ID = " & pIGID
'    Set oUtil = Nothing
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "a_Customer_P.RemoveIG(pTPID,pIGID)", Array(pTPID, pIGID)
'End Function

Public Function CustomerIndexClashes(lngTPID As Long, strAcno As String) As Boolean
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
    par.Value = strAcno
    Set par = cmd.CreateParameter("@R", adInteger, adParamOutput)
    cmd.Parameters.Append par
    par.Value = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    CustomerIndexClashes = cmd.Parameters(2)
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function


End Function


