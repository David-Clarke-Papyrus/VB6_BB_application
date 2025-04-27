VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "a_Exchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngTotalPayments As Long
Dim lngLPLessDiscExt_tot As Long
Dim lngLDeposit_tot As Long
Dim lngVAT_tot As Long
Dim lngDisc_tot As Long
Dim lngTotalQty As Long
Dim lngNominalSaleValue As Long
Dim lngTotalExtras As Long
Dim iExtrasCount As Integer
Event Valid(pMsg As String)
Event CurrRowStatus(pMsg As String)
Event Dirty(pVal As Boolean)
Event Recalculate(Redisplay As Boolean)
Event GetCustomer()
Event CreditLimitExceeded(Excess As String)
Private mExchangeCreditValue As Long
Dim mExcess As Double
Private mudtProps As ExchangeProps
Private mcolStack As Collection
Private WithEvents mColSaleLines As ch_sale
Attribute mColSaleLines.VB_VarHelpID = -1
Private mColPaymentLines As ch_Payment
Private WithEvents mCustomer As a_Customer
Attribute mCustomer.VB_VarHelpID = -1
Private colClassErrors As Collection
Private WithEvents mobjValid As z_BrokenRules
Attribute mobjValid.VB_VarHelpID = -1
Event ContainsLines(pYesNo As Boolean)
Dim lngLoyaltyValue As Long
Dim bMustPrintCNasCHange As Boolean
Dim bIDCust As Boolean
Dim lngChangeVoucherValue As Long
Dim Res As Boolean

Public Property Get ChangeVoucherValue() As Long
    ChangeVoucherValue = lngChangeVoucherValue
End Property
Public Property Get ChangeVoucherValueF() As String
    ChangeVoucherValueF = Format((ChangeVoucherValue / oPC.CurrencyDivisor), oPC.CurrencyFormat)
End Property
Public Function IssueCreditNoteForChange(pMsg As String) As Boolean
Dim oPay As a_Payment
Dim lngTotalVandCNValue As Long
Dim str As String

    lngTotalVandCNValue = 0
    For Each oPay In Me.PaymentLines
        If oPay.PaymentType = "CN" Or oPay.PaymentType = "V" Then
            lngTotalVandCNValue = lngTotalVandCNValue + oPay.Amt
        End If
    Next
    
   
    If lngTotalVandCNValue > Me.TotalPayable + oPC.CreditNoteLimit And lngTotalVandCNValue > 0 Then
        If oPC.UsageContext = "BB" Then
            str = "The total value tendered using vouchers exceeds" & Format$(oPC.CreditNoteLimit / oPC.CurrencyDivisor, oPC.CurrencyFormat$) _
                & ". Vouchers should be supplied to the customer in the next transaction." & vbCrLf _
                & "Scan vouchers in the next transaction and tender 'C' for 'cash'. " & vbCrLf
                pMsg = str
                IssueCreditNoteForChange = True
                lngChangeVoucherValue = lngTotalVandCNValue - Me.TotalPayable
        Else
            str = "The total value tendered using vouchers exceeds" & Format$(oPC.CreditNoteLimit / oPC.CurrencyDivisor, oPC.CurrencyFormat$) _
                 & " Change should be supplied as a CREDIT VOUCHER to the customer." & vbCrLf _
                 & "Please confirm." & vbCrLf & vbCrLf _
                 & "You can override this by clicking on the Override button on this form"
                 pMsg = str
                 IssueCreditNoteForChange = True
                 lngChangeVoucherValue = lngTotalVandCNValue - Me.TotalPayable
        End If
    Else
        IssueCreditNoteForChange = False
        pMsg = vbNullString
    End If
    
    
    
 '   The total value tendered using vouchers exceeds R20. Vouchers should be supplied to the customer in the next transaction.
    
    
    
End Function

Public Property Get CashTransaction() As Boolean
Dim bCT As Boolean
Dim oPay As a_Payment

    bCT = False
    For Each oPay In Me.PaymentLines
       ' If oPay.PaymentType = "C" Or oPay.PaymentType = "A" Or (Me.ChangeGiven <> 0) Then
        If oPay.PaymentType = "C" Or (Me.ChangeGiven <> 0) Then
            bCT = True
            Exit For
        End If
    Next
    CashTransaction = bCT
End Property
Private Sub mColSaleLines_ContainsLines(pYesNo As Boolean)
    RaiseEvent ContainsLines(pYesNo)
End Sub
Public Property Get OPSID() As String
    OPSID = FNS(mudtProps.OPSID)
End Property
Public Property Get ZID() As String
    ZID = FNS(mudtProps.ZID)
End Property

Public Sub StartExchange()
    mudtProps.ExchangeDate = Now()
End Sub
Public Function PaymentsComplete(Optional pChange As Long, Optional msg As String) As Boolean
    On Error GoTo errHandler
Dim oPay As a_Payment

    If mExcess > 0 Then
        PaymentsComplete = False
        Exit Function
    End If
    bMustPrintCNasCHange = False
    pChange = TotalPayment - TotalPayable
    
    msg = CheckChange(TotalPayable / oPC.CurrencyDivisor, TotalPayment / oPC.CurrencyDivisor)
    
    PaymentsComplete = (pChange >= 0)
    
    If (pChange >= 0) Then
        If PaymentLines.TotalVoucherValue > (TotalPayable + 1000) And Me.transactionType <> "R" Then 'max 1000 allowed as cash change
'''''''            Set oPay = PaymentLines.Add
'''''''            oPay.BeginEdit
'''''''            oPay.SetType "CNR"
'''''''            oPay.Amt = pChange * -1
'''''''            oPay.ApplyEdit
            bMustPrintCNasCHange = True
        End If
    End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.PaymentsComplete(pChange,msg)", Array(pChange, msg)
End Function
Private Function CheckChange(plngTotalPayable As Long, plngTotalPayment As Long) As String
    On Error GoTo errHandler
Dim i As Integer

    For i = 1 To UBound(oPC.CurrencyNotes)
        If plngTotalPayable < oPC.CurrencyNotes(i) Then
            If plngTotalPayment > oPC.CurrencyNotes(i) Then
               '' CheckChange = "NOTOK"
                Exit For
            End If
        End If
    Next i
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.CheckChange(plngTotalPayable,plngTotalPayment)", Array(plngTotalPayable, _
         plngTotalPayment)
End Function


Public Property Get MustPrintCNasChange() As Boolean
    MustPrintCNasChange = bMustPrintCNasCHange
End Property
Public Property Let TotalDiscount(val As Long)
    If mcolStack.Count = 0 Then Err.Raise 383
    SetDirty (mudtProps.TotalDiscount <> val)
    mudtProps.TotalDiscount = val
End Property

Public Property Get DiscountRate() As Double
    DiscountRate = FNDBL(mudtProps.DiscountRate)
End Property
Public Function SetDiscountRate(val As String) As Boolean
    On Error GoTo errHandler
Dim bOK As Boolean
Dim dblTEMP As Double
    dblTEMP = mudtProps.DiscountRate
    bOK = SetField_DOUBLE(mudtProps.DiscountRate, val, "DISCOUNTRATE", mcolStack.Count)
    If bOK And (dblTEMP <> mudtProps.DiscountRate) Then 'only validate if conversion is OK and the new value is other than the original
        bOK = ValidateObject("DISCOUNTRATE")
    End If
    SetDiscountRate = bOK
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SetDiscountRate(val)", val
End Function
Public Function DiscountRateF() As String
    If mudtProps.DiscountRate = 0 Then
        DiscountRateF = ""
    Else
        DiscountRateF = PBKSPercentF(mudtProps.DiscountRate)
    End If
End Function
Public Function RecalculateAllLines()
    On Error GoTo errHandler
Dim oSALE As a_Sale
    For Each oSALE In SaleLines
        oSALE.CalculateLine
    Next
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.RecalculateAllLines"
End Function
Public Function CalculateTotals(Optional Redisplay As Boolean) As Boolean
    On Error GoTo errHandler
    Res = CalculateSalesTotal()
    If Res = False Then
        CalculateTotals = False
        Exit Function
    End If
    CalculatePayments
    mudtProps.BalanceOwing = mudtProps.TotalPayable - mudtProps.TotalPayment
    If Me.transactionType <> "PA" Then
        mudtProps.ChangeGiven = mudtProps.TotalPayment - mudtProps.TotalPayable
    Else
        mudtProps.ChangeGiven = 0
    End If
    RaiseEvent Recalculate(Redisplay)
    CalculateTotals = True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.CalculateTotals"
End Function
Private Function CalculateSalesTotal() As Boolean
10        On Error GoTo errHandler
      Dim oSALE As a_Sale
      Dim oPayment As a_Payment
      Dim btmpIDCust As Boolean
      Dim cnt105 As Long
      Dim arGrp() As String
      Dim i, j, k As Long
      Dim tmpSale As a_Sale
      Dim mMBPrice As Long
      Dim mMBQty As Long
      Dim mMBCode As String
      Dim mMBDescription As String

20        lngLPLessDiscExt_tot = 0
30        lngLDeposit_tot = 0
40        lngVAT_tot = 0
50        lngDisc_tot = 0
60        iExtrasCount = 0
70        lngTotalExtras = 0
80        lngLoyaltyValue = 0
90        lngTotalQty = 0
100       lngNominalSaleValue = 0
110       bIDCust = False
120       If oPC.MultibuyRS.State <> 0 Then
130           If oPC.MultibuyRS.RecordCount > 0 Then
                  'just storing the old values already looked up in case we have to restore them
140               For Each oSALE In mColSaleLines
150                   If (Not oSALE.IsDeleted) Then
                       '   lngNominalSaleValue = lngNominalSaleValue + (oSALE.PNominalLineValue)
160                       If oSALE.TempPrice > 0 Then
170                           oSALE.Price = oSALE.TempPrice
180                       End If
190                   End If
200               Next
210           End If
220           oPC.MultibuyRS.Filter = ""
230           If oPC.MultibuyRS.RecordCount > 0 Then 'If multibuys are not applicable then skip
240               oPC.MultibuyRS.MoveFirst
250               Do While Not oPC.MultibuyRS.EOF
260                   mMBDescription = oPC.MultibuyRS.Fields("MB_Description")
270                   mMBCode = oPC.MultibuyRS.Fields("MB_SystemCode")
280                   mMBPrice = oPC.MultibuyRS.Fields("MB_PRICE")
290                   mMBQty = oPC.MultibuyRS.Fields("MB_QTYGROUP")
300                   cnt105 = 0
310                   i = 0
320                   For k = 1 To mColSaleLines.Count
330                       Set oSALE = mColSaleLines(k)
340                       If Not oSALE Is Nothing Then
350                           If UCase(oSALE.MBCode) = UCase(mMBCode) Then
360                               i = i + 1
370                               ReDim Preserve arGrp(i)
380                               arGrp(i) = oSALE.Key
390                               cnt105 = cnt105 + oSALE.Qty
400                               If cnt105 = mMBQty Then
410                                   For j = 1 To UBound(arGrp, 1)
420                                       If Not mColSaleLines.IsEditing Then mColSaleLines.BeginEdit
430                                       mColSaleLines(arGrp(j)).TempPrice = mColSaleLines(arGrp(j)).Price
440                                       mColSaleLines(arGrp(j)).Price = mMBPrice
450                                       mColSaleLines(arGrp(j)).DiscountDescription = mMBCode
460                                       mColSaleLines(arGrp(j)).DiscountRate = 0  'We do not allow discounts where multibuys are in effect
470                                       mColSaleLines(arGrp(j)).NoDiscountAllowed = True  'We do not allow discounts where multibuys are in effect
480                                       mColSaleLines.ApplyEdit
490                                       mColSaleLines.BeginEdit
500                                   Next
510                                   ReDim arGrp(0)
520                                   cnt105 = 0
530                                   i = 0
540                               End If
550                           End If
560                       End If
570                   Next
580                   oPC.MultibuyRS.MoveNext
590               Loop
600           End If
610       End If
          
620       For Each oSALE In mColSaleLines
630           If (Not oSALE.IsDeleted) Then
640               lngNominalSaleValue = lngNominalSaleValue + (oSALE.PNominalLineValue)
650           End If
660       Next
670       For Each oSALE In mColSaleLines
680           If (Not oSALE.IsDeleted) Then
           '       lngNominalSaleValue = lngNominalSaleValue + (oSALE.PNominalLineValue)
690               If oSALE.MBCode = "" Then
700                   oSALE.FindRule lngNominalSaleValue, btmpIDCust
710               End If
720               If btmpIDCust Then bIDCust = True
730               Res = oSALE.CalculateLine()
                    If Res = False Then
                        CalculateSalesTotal = False
                        Exit Function
                    End If
740               If oSALE.BottomOfDocument Then
750                   iExtrasCount = iExtrasCount + 1
760               End If
770               lngTotalQty = lngTotalQty + oSALE.Qty
780               lngVAT_tot = lngVAT_tot + oSALE.PLessDiscExtVAT()
790               lngDisc_tot = lngDisc_tot + oSALE.PDiscExt()
800               lngLPLessDiscExt_tot = lngLPLessDiscExt_tot + oSALE.PLessDiscExt()
810               If Not Me.Customer Is Nothing Then
820                   If oSALE.DiscountRate = 0 And Customer.IsLoyaltyCustomer Then     'only loyalty value on non-discounted items
830                       If oSALE.LoyaltyDiscount > oPC.LoyaltyRate And oSALE.NonStock = False Then   'we dont give loyalty value on newspapers etc
840                           lngLoyaltyValue = lngLoyaltyValue + (oSALE.PLessDiscExt() * (oSALE.LoyaltyDiscount / 100))
850                       ElseIf oSALE.NonStock = False And oSALE.NoDiscountAllowed = False Then
860                           lngLoyaltyValue = lngLoyaltyValue + (oSALE.PLessDiscExt() * (oPC.LoyaltyRate / 100))
870                       End If
880                   End If
890               End If
900               If oSALE.BottomOfDocument Then
910                   lngTotalExtras = lngTotalExtras + oSALE.PLessDiscExt()
920               End If
930           End If
940       Next
950       If bIDCust Then RaiseEvent GetCustomer
960       lngTotalPayments = 0
970       For Each oPayment In mColPaymentLines
980           If Not (oPayment.PaymentType = "CN" And oPayment.Amt < 0) Then
990               lngTotalPayments = lngTotalPayments + oPayment.Amt
1000          End If
1010      Next
1020      mudtProps.TotalQty = lngTotalQty
1030      mudtProps.TotalExtras = lngTotalExtras
1040      mudtProps.TotalDiscount = lngDisc_tot
1050      mudtProps.TotalVAT = lngVAT_tot
1060      mudtProps.TotalExtension = lngLPLessDiscExt_tot
1070      mudtProps.TotalExVAT = mudtProps.TotalExtension - mudtProps.TotalVAT
1080      If FNS(mudtProps.Type) <> "OR" Then
1090          mudtProps.TotalPayable = lngLPLessDiscExt_tot - lngLDeposit_tot
1100      Else
1110          mudtProps.TotalPayable = lngLPLessDiscExt_tot
1120      End If
1130      If (Not VATable) And ShowVAT And (FNS(mudtProps.Type) <> "OR") Then
1140          mudtProps.TotalPayable = lngLPLessDiscExt_tot - lngVAT_tot
1150      End If
          
          'This has been removed 19/7/2010 Seems to make the totals on the screen inconsistent
         ' mudtProps.TotalPayable = RoundDown(mudtProps.TotalPayable, oPC.GetRoundingRule(mudtProps.TotalPayable))
          
          
1160      If mudtProps.TotalPayable >= 10000 And Me.Customer.CustomerType = oPC.GetLoyaltyCode Then
1170          mudtProps.LoyaltyValue = lngLoyaltyValue
1180      Else
1190          mudtProps.LoyaltyValue = 0
1200      End If

          CalculateSalesTotal = True
1210      Exit Function

errHandler:
          ErrPreserve
          CalculateSalesTotal = False
          Exit Function
1220      If ErrMustStop Then Debug.Assert False: Resume
1230      ErrorIn "a_Exchange.CalculateSalesTotal", , , , "line number", Array(Erl())
End Function
Public Property Get NeedsCustomerInfo() As Boolean
    If (TotalPayable < 0 And Me.Customer.CustomerID = 0) Or (Me.PaymentLines.NeedsCustomerInfo = True And Me.Customer.CustomerID = 0) Then
        NeedsCustomerInfo = True
    Else
        NeedsCustomerInfo = False
    End If
End Property
Public Property Get CustomerToBeCredited() As Boolean
    If (transactionType = "S" And TotalPayable < 0) Then
        CustomerToBeCredited = True
    Else
        CustomerToBeCredited = False
    End If
End Property

Public Property Get BalanceOwing() As Long
    BalanceOwing = mudtProps.BalanceOwing
End Property
Public Property Get BalanceOwingF() As String
    BalanceOwingF = Format$((BalanceOwing / oPC.CurrencyDivisor), oPC.CurrencyFormat$)
End Property
Public Function CanCancel(pExchangeNumber As Long)
    On Error GoTo errHandler
Dim iResult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    oPC.OpenLocalDatabase
    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_CanCancel"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@EXCHANGENUMBER", adInteger, adParamInput, , pExchangeNumber)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("RES", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.DBLocalConn
    cmd.Execute
    
    CanCancel = (cmd.Parameters(1) = 0)
    Set cmd = Nothing

    oPC.CloseLocalDatabase
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.CanCancel(pEXCHANGENUMBER)", pExchangeNumber
End Function
Private Function CalculatePayments()
    On Error GoTo errHandler
Dim oPayment As a_Payment
Dim bLineDiscountApplied As Boolean

    lngTotalPayments = 0
    mExchangeCreditValue = 0
    mExcess = 0
    RaiseEvent CreditLimitExceeded("")
    For Each oPayment In mColPaymentLines
        lngTotalPayments = lngTotalPayments + oPayment.Amt
        If oPayment.PaymentType = "AC" Then
            mExchangeCreditValue = (mExchangeCreditValue + oPayment.Amt) / oPC.CurrencyDivisor
            mExcess = (Me.Customer.Balance + mExchangeCreditValue) - (Me.Customer.CreditLimit)
            If mExcess > 0 Then
                RaiseEvent CreditLimitExceeded(Format((mExcess), oPC.CurrencyFormat))
            End If
        End If
    Next
    mudtProps.TotalPayment = lngTotalPayments
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.CalculatePayments"
End Function
Public Property Get CreditLimitExceededBy() As Double
    CreditLimitExceededBy = mExcess
End Property
Public Property Get CreditLimitExceededByF() As String
    CreditLimitExceededByF = Format((CreditLimitExceededBy), oPC.CurrencyFormat)
End Property


Public Property Get ExchangeNumber() As Long
    ExchangeNumber = mudtProps.ExchangeNumber
End Property

Public Property Get TotalQty() As Long
    TotalQty = lngTotalQty
End Property
Public Property Get NominalValue() As Long
    NominalValue = lngNominalSaleValue
End Property
Public Property Get NominalValueF() As String
    NominalValueF = Format$((lngNominalSaleValue / oPC.CurrencyDivisor), oPC.CurrencyFormat$)
End Property
Public Property Get LoyaltyValue() As Long
    On Error GoTo errHandler
        LoyaltyValue = mudtProps.LoyaltyValue
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.LoyaltyValue()"
End Property
Public Property Get LoyaltyValueF() As String
    On Error GoTo errHandler
        LoyaltyValueF = Format$((LoyaltyValue / oPC.CurrencyDivisor), oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.LoyaltyValueF()"
End Property



Public Property Get TotalExtWithOutExtrasF() As String
    On Error GoTo errHandler
        TotalExtWithOutExtrasF = Format$((mudtProps.TotalExtension - mudtProps.TotalExtras) / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalExtWithOutExtrasF()"
End Property
Public Property Get TotalExVATF() As String
    On Error GoTo errHandler
        TotalExVATF = Format$((mudtProps.TotalExVAT) / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalExVATF()"
End Property

Public Property Get TotalLessDiscExtF() As String
    On Error GoTo errHandler
        TotalLessDiscExtF = Format$(mudtProps.TotalExtension / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalLessDiscExtF()"
End Property
Public Property Get TotalLessDiscExt() As String
    On Error GoTo errHandler
        TotalLessDiscExt = mudtProps.TotalExtension
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalLessDiscExt()"
End Property
Public Property Get TotalVATF() As String
    On Error GoTo errHandler
        TotalVATF = Format$(mudtProps.TotalVAT / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalVATF()"
End Property
Public Property Let TotalPayable(val As Long)
        mudtProps.TotalPayable = val
End Property

Public Property Get TotalPayable() As Long
    On Error GoTo errHandler
        TotalPayable = mudtProps.TotalPayable
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalPayable()"
End Property
Public Property Get TotalPayableF() As String
    On Error GoTo errHandler
    Dim tmp As Long
    
 '   If mudtProps.TotalPayable < 0 Then
 '       tmp = 0
 '   Else
        tmp = mudtProps.TotalPayable
 '   End If
       TotalPayableF = Format$(tmp / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalPayableF()"
End Property
Public Property Get TotalDiscount() As Long
    On Error GoTo errHandler
        TotalDiscount = mudtProps.TotalDiscount
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalDiscount()"
End Property
Public Property Get TotalDiscountF() As String
    On Error GoTo errHandler
        TotalDiscountF = Format$(mudtProps.TotalDiscount / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalDiscountF()"
End Property
Public Property Get TotalDepositF() As String
    On Error GoTo errHandler
        TotalDepositF = Format$(TotalDeposit(), oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalDepositF()"
End Property
Public Property Get TotalDeposit() As Double
    On Error GoTo errHandler
        TotalDeposit = lngLDeposit_tot / oPC.CurrencyDivisor
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalDeposit()"
End Property
Public Property Get TotalPayment() As Long
    On Error GoTo errHandler
        TotalPayment = mudtProps.TotalPayment
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalPayment()"
End Property
Public Property Get TotalPaymentF() As String
    On Error GoTo errHandler
        TotalPaymentF = Format$(mudtProps.TotalPayment / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TotalPaymentF()"
End Property
Public Property Let TotalVAT(val As Long)
    mudtProps.TotalVAT = val
    SetDirty True
End Property
Private Function GetState() As String
    On Error GoTo errHandler
  Dim udtData As ExchangeData
  LSet udtData = mudtProps
  GetState = udtData.buffer
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.GetState"
End Function
Private Sub SetState(buffer As String)
    On Error GoTo errHandler
  Dim udtData As ExchangeData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SetState(buffer)", buffer
End Sub
Public Function GetSuperState() As String
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  
  Set objPB = New PropertyBag
  With objPB
    .WriteProperty "State", GetState
    .WriteProperty "SALES", mColSaleLines.GetSuperState
    .WriteProperty "PAYMENTS", mColPaymentLines.GetSuperState
    GetSuperState = .Contents
  End With
  Set objPB = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.GetSuperState"
End Function
Public Sub SetSuperState(buffer As String)
    On Error GoTo errHandler
  Dim objPB As PropertyBag
  Dim arBuffer() As Byte
  
  Set objPB = New PropertyBag
  arBuffer = buffer
  With objPB
    .Contents = arBuffer
    SetState .ReadProperty("State")
    mColSaleLines.SetSuperState .ReadProperty("SALES")
    mColPaymentLines.SetSuperState .ReadProperty("PAYMENTS")
  End With
  Set objPB = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SetSuperState(buffer)", buffer
End Sub


Public Sub BeginEdit()
    On Error GoTo errHandler
    mColSaleLines.BeginEdit
    mColPaymentLines.BeginEdit
    mcolStack.Add GetState
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.BeginEdit"
End Sub

Public Sub CancelEdit()
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 445
    mColSaleLines.CancelEdit
    mColPaymentLines.CancelEdit
    mudtProps.IsDeleted = False
    With mcolStack
      SetState .Item(.Count)
      .Remove .Count
    End With
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.CancelEdit"
End Sub
Public Sub PostExchange()
    On Error GoTo errHandler
Dim oSM As z_SM
Dim oPay As a_Payment
Dim oSALELine As a_Sale

    'If payment type is a redeemed deposit then we must set the deposit status to 'E' - used in Exchange
    Set oSM = New z_SM
    Select Case Me.transactionType
    Case "S"       'Sale
        If Me.ToVoid > 0 Then
            oSM.CancelRedeemedDeposit Me.ToVoid
        End If
        For Each oPay In Me.PaymentLines
            If oPay.PaymentType = "RD" Then
                oSM.MarkRedeemedDeposit oPay.COLID
            End If
        Next
    Case "R"    'Return
    Case "C"
    Case "PC"  'Petty cash withdrawal
    Case "PCC"  'Petty cash credit
    Case "APP"
    Case "AR"
    Case "DEP"    'Deposit accepted
        For Each oSALELine In Me.SaleLines
            If oSALELine.COLID > 0 Then
                oSM.PayDeposit oSALELine.COLID
            End If
        Next
    Case "RDEP"    'Deposit refunded
        For Each oSALELine In Me.SaleLines
            If oSALELine.COLID > 0 Then
                oSM.RefundDeposit oSALELine.COLID
            End If
        Next
    Case "OR"
    Case "V"
        If Me.ToVoid > 0 Then
            oSM.CancelRedeemedDeposit Me.ToVoid
        End If

    Case "X"
    Case "Z"
    Case "OD"
    Case "PA"
    Case "CN"
    Case "A"
    Case Else
        MsgBox "Invalid transaction type in PostExchange: " & Me.transactionType
        
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.PostExchange"
End Sub

Public Function ApplyEdit(Optional pMsg As String) As String

10        On Error GoTo errHandler
      Dim strTemp As String
      Dim objPersist As a_Exchange_P
      Dim strPos As String
          
30        ApplyEdit = ""
50        Res = CalculateTotals()
          If Res = False Then
            ApplyEdit = "calc"
            Exit Function
          End If
60        mudtProps.TPID = Me.Customer.ID
70        mudtProps.ExchangeDate = Now()
90        mudtProps.ZID = oPC.ZSession.Current_Z_Session_ID
100       mudtProps.OPSID = oPC.ZSession.OpSession.OPSID
110       mudtProps.ExchangeNumber = oPC.ExchangeNumber
120       mudtProps.TPID = Me.Customer.ID
140       If mcolStack.Count = 0 Then Err.Raise 445
150       mColSaleLines.ApplyEdit
170       mColPaymentLines.ApplyEdit
180       Set objPersist = New a_Exchange_P
          
200       If mudtProps.IsDeleted Then
210           objPersist.DeleteObject mudtProps.ID
220           mcolStack.Remove mcolStack.Count
230           mudtProps.IsNew = True
240           mudtProps.IsDeleted = False
260       ElseIf IsDirty Or mudtProps.IsNew Then
270           If Not IsValid Then Err.Raise 445
280           strTemp = objPersist.Save(GetSuperState)
          
300           If UCase(Left(strTemp, 7)) <> "TIMEOUT" Then
310               mcolStack.Remove mcolStack.Count
320               SetSuperState strTemp
330               mudtProps.IsNew = False
340           Else 'save failed
350               mColSaleLines.BeginEdit
360               mColPaymentLines.BeginEdit
370               If Not IsMissing(pMsg) Then
380                   pMsg = strTemp
390               End If
400               ApplyEdit = "TIMEOUT"
410               Set objPersist = Nothing
420               GoTo EXIT_Handler
430           End If
          
450       Else
460           mcolStack.Remove mcolStack.Count
470       End If
490       Set objPersist = Nothing
500       mudtProps.IsDirty = False
          
EXIT_Handler:
510       Exit Function
errHandler:
520       If ErrMustStop Then Debug.Assert False: Resume
530       ErrorIn "a_Exchange.ApplyEdit(pMsg)", pMsg, , , "Line number", Array(Erl())


End Function
Public Property Get IsDeleted() As Boolean
    On Error GoTo errHandler
    IsDeleted = mudtProps.IsDeleted
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.IsDeleted"
End Property
Public Property Get IsEditing() As Boolean
    On Error GoTo errHandler
    IsEditing = mcolStack.Count > 0
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.IsEditing"
End Property

Public Property Get IsNew() As Boolean
    On Error GoTo errHandler
    IsNew = mudtProps.IsNew
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.IsNew"
End Property

Public Property Get IsDirty() As Boolean
    On Error GoTo errHandler
    IsDirty = mudtProps.IsDirty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.IsDirty"
End Property

Public Property Get IsValid() As Boolean
    On Error GoTo errHandler
    IsValid = (mobjValid.Count = 0)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.IsValid"
End Property
Private Sub Class_Initialize()
    On Error GoTo errHandler
    Set mcolStack = New Collection
    Set mobjValid = New z_BrokenRules
    Set mColSaleLines = New ch_sale
    mColSaleLines.component Me
    Set mColPaymentLines = New ch_Payment
    mColPaymentLines.component Me
    Set mCustomer = New a_Customer
    mudtProps.IsNew = True
    mudtProps.ShowVAT = True
    mudtProps.Status = "S"
    mudtProps.OPSID = oPC.ZSession.OpSession.OperatorID
    mudtProps.ExchangeNumber = oPC.ExchangeNumber
    LoadClassErrorsCollection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Class_Initialize"
End Sub

Private Sub Class_Terminate()
    On Error GoTo errHandler
'    If mcolStack.Count > 0 Then _
'      Err.Raise vbObjectError + 1001, , "State stack is not empty"
      
    Set mcolStack = Nothing
    Set mobjValid = Nothing
    Set mColSaleLines = Nothing
    Set mColPaymentLines = Nothing
    Set mCustomer = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Class_Terminate"
End Sub

Private Sub mobjValid_BrokenRule(errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.mobjValid_BrokenRule(errors)", errors
End Sub
Private Sub mobjValid_RuleUnbroken(errors As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(errors))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.mobjValid_RuleUnbroken(errors)", errors
End Sub

Private Sub mobjValid_NoBrokenRules()
    On Error GoTo errHandler
    RaiseEvent Valid("")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.mobjValid_NoBrokenRules"
End Sub
Private Sub mobjValid_Status(pMsg As String)
    On Error GoTo errHandler
    RaiseEvent Valid(TranslateErrors(pMsg))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.mobjValid_Status(pMsg)", pMsg
End Sub

Public Sub Load(pEXCHID As String, ReadOnly As Boolean)
    On Error GoTo errHandler
Dim objPersist As a_Exchange_P
    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    mudtProps.IsNew = False
    
    Set objPersist = New a_Exchange_P
    SetSuperState objPersist.Fetch(pEXCHID)
    mCustomer.Load mudtProps.TPID
    Set objPersist = Nothing
    mobjValid.RuleBroken "TP", False
    If Me.transactionType <> "OR" Then
        Res = CalculateTotals()
        If Res = False Then
            MsgBox "Problem calculating this transaction", vbInformation + vbOKOnly, "Status"
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Load(pEXCHID,ReadOnly)", Array(pEXCHID, ReadOnly)
End Sub
Public Sub LoadFromMainDB(pEXCHID As String, ReadOnly As Boolean, Optional Successful As Boolean = True)
    On Error GoTo errHandler
Dim objPersist As a_Exchange_P
Dim buf As String

'    If mcolStack.Count > 0 Then Err.Raise 445
    If Not mudtProps.IsNew Then Err.Raise 445
    
    Successful = False
    oPC.dbConnectMain
    If oPC.IsConnectedToServer = False Then Exit Sub
    mudtProps.IsNew = False
    
    Set objPersist = New a_Exchange_P
    buf = objPersist.FetchFromMainDB(pEXCHID)
    If buf = "" Then Exit Sub
    
    SetSuperState buf
    mCustomer.Load mudtProps.TPID
    Set objPersist = Nothing
    
    oPC.dbMainDisConnect
    
    mobjValid.RuleBroken "TP", False
    If Me.transactionType <> "OR" Then
        Res = CalculateSalesTotal()
        If Res = False Then
            Successful = False
            Exit Sub
        End If
        CalculatePayments
    End If
    mudtProps.BalanceOwing = mudtProps.TotalPayable - mudtProps.TotalPayment
    If Me.transactionType <> "PA" Then
        mudtProps.ChangeGiven = mudtProps.TotalPayment - mudtProps.TotalPayable
    Else
        mudtProps.ChangeGiven = 0
    End If
    Successful = True
    Exit Sub
errHandler:
    Successful = False
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Load(pEXCHID,ReadOnly)", Array(pEXCHID, ReadOnly)
End Sub

Public Property Get Customer() As a_Customer
    On Error GoTo errHandler
    Set Customer = mCustomer
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Customer"
End Property
Public Sub Delete()
    On Error GoTo errHandler
  If mcolStack.Count = 0 Then Err.Raise 445
  
  mudtProps.IsDeleted = True
  SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Delete"
End Sub

Public Property Get ExchangeID() As String
    On Error GoTo errHandler
    ExchangeID = FNS(mudtProps.ID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ExchangeID"
End Property

Public Property Let ExchangeID(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ExchangeID(val)", val
End Property

Public Function SetGeneralDiscount(val As String) As Boolean
    On Error GoTo errHandler
Dim dblDiscountRate As Double
Dim oSALE As a_Sale

    If mcolStack.Count = 0 Then Err.Raise 383
    SetGeneralDiscount = True
    If Trim$(val) = "" Then
        dblDiscountRate = 0
    ElseIf Not ConvertToDBL(val, dblDiscountRate) Then
        SetGeneralDiscount = False
        Exit Function
    End If
    If (dblDiscountRate <= 0) Then
        mobjValid.RuleBroken "DISCOUNT", True
    Else
        mobjValid.RuleBroken "DISCOUNT", False
    End If
    mudtProps.DiscountRate = dblDiscountRate
    SetDirty True
    For Each oSALE In SaleLines  'disctribute discount rate over all lines where non nonstaock items exists
        If Not oSALE.NonStock Then
            oSALE.DiscountRate = dblDiscountRate
        End If
    Next
    RecalculateAllLines
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SetGeneralDiscount(val)", val
End Function
Public Function SetCustomer(pTPID As Long) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean
    If pTPID = 0 Then
        Set mCustomer = Nothing
        Exit Function
    End If
    bSuccess = Customer.Load(pTPID)
    SetCustomer = bSuccess
    If bSuccess Then
        VATable = Customer.VATable
        mobjValid.RuleBroken "TP", False
        SetDirty True
    End If
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SetCustomer(pTPID)", pTPID
End Function
Public Function SetCustomerFromAccNum(pAccnum As String) As Boolean
    On Error GoTo errHandler
Dim bSuccess As Boolean

    bSuccess = Me.Customer.Load(, pAccnum)
    SetCustomerFromAccNum = bSuccess
    If bSuccess Then
        mobjValid.RuleBroken "TP", False
    End If
    Me.VATable = Customer.VATable
    SetDirty True
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SetCustomerFromAccNum(pAccnum)", pAccnum
End Function
Public Property Let ShowVAT(val As Boolean)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ShowVAT = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ShowVAT(val)", val
End Property
Public Property Get ShowVAT() As Boolean
    On Error GoTo errHandler
    ShowVAT = mudtProps.ShowVAT
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ShowVAT"
End Property
Public Property Get SupervisorID() As Long
    On Error GoTo errHandler
    SupervisorID = mudtProps.SupervisorID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SupervisorID"
End Property
Public Property Let SupervisorID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.SupervisorID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SupervisorID(val)", val
End Property
Public Property Get SalesPersonName() As String
    SalesPersonName = FNS(mudtProps.SalesPersonName)
End Property
Public Property Let SalesPersonName(val As String)
    mudtProps.SalesPersonName = val
End Property
Public Property Get Note() As String
    On Error GoTo errHandler
    Note = FNS(mudtProps.Note)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.CustomerName"
End Property


Public Property Let Note(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.Note = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Note(val)", val
End Property

Public Property Get OperatorID() As Long
    On Error GoTo errHandler
    OperatorID = mudtProps.OperatorID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.OperatorID"
End Property
Public Property Let OperatorID(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.OperatorID = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.OperatorID(val)", val
End Property
Public Property Get ToVoid() As Long
    On Error GoTo errHandler
    ToVoid = mudtProps.ToVoid
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.OperatorID"
End Property
Public Property Let ToVoid(val As Long)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ToVoid = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ToVoid(val)", val
End Property

Public Property Get StaffName() As String
    On Error GoTo errHandler
    StaffName = FNS(mudtProps.StaffName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.StaffName"
End Property
Public Property Let StaffName(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.StaffName = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.StaffName(val)", val
End Property
Public Property Get ChangeGiven() As Long
    On Error GoTo errHandler
    ChangeGiven = FNN(mudtProps.ChangeGiven)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ChangeGiven"
End Property
Public Property Get ChangeGivenF() As String
    On Error GoTo errHandler
    ChangeGivenF = Format$(FNN(mudtProps.ChangeGiven) / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ChangeGivenF"
End Property
Public Property Get ChangeGivenNonNegativeF() As String
    On Error GoTo errHandler
    ChangeGivenNonNegativeF = Format$(IIf(FNN(mudtProps.ChangeGiven) < 0, FNN(mudtProps.ChangeGiven) * -1, FNN(mudtProps.ChangeGiven)) / oPC.CurrencyDivisor, oPC.CurrencyFormat$)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ChangeGivenF"
End Property

Public Property Get CCPayment() As a_Payment
Dim oPayment As a_Payment

    For Each oPayment In Me.PaymentLines
        If oPayment.PaymentType = "A" Then
            Set CCPayment = oPayment
            Exit For
        End If
    Next

End Property


Public Property Get ExchangeCode() As String
    On Error GoTo errHandler
    ExchangeCode = FNS(mudtProps.ExchangeCode)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.DocCode"
End Property
Public Property Let ExchangeCode(val As String)
    On Error GoTo errHandler
    If mcolStack.Count = 0 Then Err.Raise 383
    mudtProps.ExchangeCode = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ExchangeCode(val)", val
End Property


Public Property Get ExchangeDate() As Date
    On Error GoTo errHandler
    ExchangeDate = mudtProps.ExchangeDate
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ExchangeDate"
End Property
Public Property Get ExchangeDateF() As String
    On Error GoTo errHandler
    ExchangeDateF = Format$(mudtProps.ExchangeDate, "dd/mm/yyyy")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ExchangeDateF"
End Property
Public Property Get ExchangeTimeF() As String
    On Error GoTo errHandler
    ExchangeTimeF = Format$(mudtProps.ExchangeDate, "HH:NN")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ExchangeDateF"
End Property
Public Property Get ExchangeDateTimeF() As String
    On Error GoTo errHandler
    ExchangeDateTimeF = Format$(mudtProps.ExchangeDate, "dd/mm/yyyy HH:NN")
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ExchangeDateTimeF"
End Property

Public Sub SetExchangeType(pType As enumExchangeType)
    If pType = eSaleType Then
        mudtProps.Type = "S"
    ElseIf pType = ereturntype Then
        mudtProps.Type = "R"
    ElseIf pType = ePettyCashType Then
        mudtProps.Type = "PC"
    ElseIf pType = eDepositType Then
        mudtProps.Type = "DEP"
    ElseIf pType = eReturnDepositType Then
        mudtProps.Type = "RDEP"
    ElseIf pType = eApproType Then
        mudtProps.Type = "APP"
    ElseIf pType = eCreditVoucherType Then
        mudtProps.Type = "C"
    ElseIf pType = ePettyCashCreditType Then
        mudtProps.Type = "PCC"
    ElseIf pType = eAccountPaymentType Then
        mudtProps.Type = "PA"
    ElseIf pType = eApproReturnType Then
        mudtProps.Type = "AR"
    ElseIf pType = eOrderRequestType Then
        mudtProps.Type = "OR"
    ElseIf pType = eVoidAction Then
        mudtProps.Type = "V"
    ElseIf pType = eXEND Then
        mudtProps.Type = "X"
    ElseIf pType = eZEND Then
        mudtProps.Type = "Z"
    ElseIf pType = eOpenDrawerType Then
        mudtProps.Type = "OD"
    ElseIf pType = eAccountCreditNoteType Then
        mudtProps.Type = "CN"
    ElseIf pType = eAccountSaleType Then
        mudtProps.Type = "A"
    End If

End Sub
'Public Property Let TotalVAT(val As Long)
'    mudtProps.TotalVAT = val
'    SetDirty True
'End Property
Public Property Get TotalVAT() As Long
    TotalVAT = mudtProps.TotalVAT
End Property
Public Property Get VATable() As Boolean
    On Error GoTo errHandler
  VATable = Me.Customer.VATable
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.VATable"
End Property
Public Property Let VATable(val As Boolean)
    On Error GoTo errHandler
    mudtProps.VATable = val
    SetDirty True
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.VATable(val)", val
End Property
Public Property Get SalesRepID() As Long
    SalesRepID = mudtProps.SalesRepID
End Property
Public Property Let SalesRepID(val As Long)
    mudtProps.SalesRepID = val
End Property

'Public Property Get DocumentHeadingDescription() As String
'Dim i As Integer
'Dim iSold As Integer
'Dim iReturned As Integer
'
'    For i = 1 To Me.SaleLines.Count
'        If Me.SaleLines(i).Qty > 0 Then
'            iSold = iSold + Me.SaleLines(i).Qty
'        Else
'            iReturned = iReturned + Me.SaleLines(i).Qty
'        End If
'    Next
'    If iReturned = 0 Then
'        DocumentHeadingDescription = ""
'End Property
Public Property Get SaleLines() As ch_sale
  Set SaleLines = mColSaleLines
End Property
Public Property Get PaymentLines() As ch_Payment
  Set PaymentLines = mColPaymentLines
End Property
Public Property Let IdentifyCustomer(val As Boolean)
    mudtProps.IdentifyCustomer = FNB(val)
End Property
Public Property Get IdentifyCustomer() As Boolean
    IdentifyCustomer = FNB(mudtProps.IdentifyCustomer)
End Property
Public Property Get StatusF() As String
    On Error GoTo errHandler
    If mudtProps.Status = stVOID Then
        StatusF = "VOID"
    ElseIf mudtProps.Status = stInProcess Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.Status = stISSUED Then
        StatusF = "ISSUED"
    ElseIf mudtProps.Status = stCOMPLETE Then
        StatusF = "COMPLETE"
    ElseIf mudtProps.Status = stPROFORMA Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.Status = stCANCELLED Then
        StatusF = "CANCELLED"
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.StatusF"
End Property
Public Property Get Status() As enStatus
    On Error GoTo errHandler
    Status = mudtProps.Status
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Status"
End Property

Public Function PrintExchange_Display(Optional ByVal pRoundedUp As Boolean) As String
    On Error GoTo errHandler
Dim bDiscountExists As Boolean
Dim oSALE As a_Sale
Dim strExchange As String
Dim i As Long
    
    strExchange = "Document No. " & ExchangeCode & vbCrLf
    For i = 1 To Me.SaleLines.Count
        Set oSALE = Me.SaleLines.FindLineBySeqNum(i)
        If Not oSALE Is Nothing Then
            If Not oSALE.IsDeleted Then
                If oSALE.BottomOfDocument = False Then
                    With oSALE
                        If .DiscountRate > 0 Then bDiscountExists = True
                        If .NonStock Then
                            strExchange = strExchange & .TitleF(25) & "  " & .PLessDiscExtF()
                        Else
                            strExchange = strExchange & .CodeF & "  " & .QtyF & "  " & .TitleF(25) & "  " & .PriceF() & "  " & IIf(.DiscountRate = 0, "", .DiscountRateF) & "  " & .PLessDiscExtF()
                        End If
                    End With
                strExchange = strExchange & vbCrLf
                Else
                    iExtrasCount = iExtrasCount + 1
                End If
            End If
        End If
    Next i
    strExchange = strExchange & vbCrLf
    If iExtrasCount > 0 Then  'then subtotal 2 is necessary
        strExchange = strExchange & "subtotal    " & Me.TotalExtWithOutExtrasF() & vbCrLf
        For i = 1 To SaleLines.Count
            If SaleLines(i).BottomOfDocument = True Then
                With SaleLines(i)
                    If .NonStock Then
                        strExchange = strExchange & .TitleF(25) & "    " & .PriceF() & vbCrLf
                    End If
                End With
            End If
        Next i
    End If
    If Me.VATable Then ''''''''''(1)
        strExchange = strExchange & "Total " & TotalPayableF() & vbCrLf
        strExchange = strExchange & "Includes VAT of " & TotalVATF() & vbCrLf
    Else  'overseas clients
        If Me.ShowVAT Then
            strExchange = strExchange & "subtotal    " & TotalLessDiscExtF() & vbCrLf   'subtotal 3
            strExchange = strExchange & "Less VAT of    " & TotalVATF() & vbCrLf
            strExchange = strExchange & "Payable    " & TotalPayableF() & vbCrLf
        Else
            strExchange = strExchange & "Total   " & TotalPayableF() & vbCrLf
        End If
    End If
    If TotalDeposit() > 0 Then
        strExchange = strExchange & "Less deposits paid: " & TotalDepositF()
    End If
    PrintExchange_Display = strExchange
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.PrintExchange_Display(pRoundedUp)", pRoundedUp
End Function



Private Function TranslateErrors(pRawErrors As String) As String
    On Error GoTo errHandler
Dim strOut As String
Dim strRule, strAllRules As String
Dim NoMoreRules As Boolean
Dim iMarker, istart As Integer
    iMarker = 1
    strAllRules = ""
    If Len(pRawErrors) > 0 Then
        iMarker = InStr(iMarker + 1, pRawErrors, ",")
        If iMarker > 0 Then
            strAllRules = colClassErrors(Left(pRawErrors, iMarker - 1))
        Else
            strAllRules = colClassErrors(pRawErrors)
        End If
        Do Until iMarker = 0
            istart = iMarker + 1
            iMarker = InStr(istart, pRawErrors, ",")
            If iMarker > 0 Then
                strRule = colClassErrors(MID(pRawErrors, istart, iMarker - istart))
            Else
                strRule = colClassErrors(MID(pRawErrors, istart))
            End If
                
            strAllRules = strAllRules & vbCrLf & strRule
        Loop
    End If
    TranslateErrors = strAllRules
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.TranslateErrors(pRawErrors)", pRawErrors
End Function
Private Sub LoadClassErrorsCollection()
    On Error GoTo errHandler
    Set colClassErrors = New Collection
    colClassErrors.Add "Missing customer", "TP"
    colClassErrors.Add "Note too short", "NOTE"
    colClassErrors.Add "Invalid discount", "DISCOUNT"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.LoadClassErrorsCollection"
End Sub
Private Function ValidateObject(pFld As String) As Boolean
    On Error GoTo errHandler
Dim bValid As Boolean
    bValid = True
    If mudtProps.TPID = 0 Then
        mobjValid.BreakRule "TP", True
        If pFld = "TP" Then bValid = False
    Else
        mobjValid.BreakRule "TP", False
    End If

    mobjValid.GetStatus
    SetDirty bValid
    ValidateObject = bValid
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.ValidateObject(pFld)", pFld
End Function
Private Sub SetDirty(pVal As Boolean)
    On Error GoTo errHandler
    mudtProps.IsDirty = pVal
    RaiseEvent Dirty(pVal)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SetDirty(pVal)", pVal
End Sub


Public Sub SetStatus(val As enStatus)
    On Error GoTo errHandler
    mudtProps.Status = val
    SetDirty True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.SetStatus(val)", val
End Sub
Public Function VoidDocument()
    On Error GoTo errHandler
Dim strresult As String
    Me.BeginEdit
    Me.SetStatus stVOID
    Me.ApplyEdit
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.VoidDocument"
End Function
Public Sub Reload()
    On Error GoTo errHandler
Dim lngTRID As String
    lngTRID = Me.ExchangeID
    Class_Terminate
    Class_Initialize
    Load lngTRID, True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "a_Exchange.Reload"
End Sub

Public Function DisplayPaymentTotals() As String
Dim str As String
Dim i As Long
    str = "Payments---------------"
    For i = 1 To PaymentLines.Count
        str = str & PaymentLines(i).PaymentTypeF & ":   " & PaymentLines(i).AmtF
    Next
    DisplayPaymentTotals = str
End Function

Public Sub PrintSlip()
'Dim oTillSlip As arTillSlip
'    Set oTillSlip = New arTillSlip
'    oTillSlip.Show vbModal
End Sub
Public Property Get transactionType() As String
    transactionType = FNS(mudtProps.Type)
End Property
Public Property Get TransactionTypeEnum() As Long
    Select Case Trim(mudtProps.Type)
    Case "S"
        TransactionTypeEnum = eSaleType
    Case "R"
        TransactionTypeEnum = ereturntype
    Case "PC"
        TransactionTypeEnum = ePettyCashType
    Case "DEP"
        TransactionTypeEnum = eDepositType
    Case "RDEP"
        TransactionTypeEnum = eReturnDepositType
    Case "APP"
        TransactionTypeEnum = eApproType
    Case "C"
        TransactionTypeEnum = eCreditVoucherType
    Case "PCC"
        TransactionTypeEnum = ePettyCashCreditType
    Case "AR"
        TransactionTypeEnum = eApproReturnType
    Case "OR"
        TransactionTypeEnum = eOrderRequestType
    Case "V"
        TransactionTypeEnum = eVoidAction
    Case "X"
        TransactionTypeEnum = eXEND
    Case "Z"
        TransactionTypeEnum = eZEND
    Case "OD"
        TransactionTypeEnum = eOpenDrawerType
    Case "PA"
        TransactionTypeEnum = eAccountPaymentType
    Case "CN"
        TransactionTypeEnum = eAccountCreditNoteType
    Case "A"
        TransactionTypeEnum = eAccountSaleType
    Case "A"
        TransactionTypeEnum = eAccountSaleType
    End Select

End Property
