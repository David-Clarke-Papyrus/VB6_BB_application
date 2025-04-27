VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Cashup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngACRec As Long
Dim lngTotalCash As Long
Dim lngTotalCheques As Long
Dim lngCOGS As Long
Dim lngTotalVouchersRedeemed As Long
Dim lngTotalVouchersEXCL As Long
Dim lngTotalVouchersINCL As Long
Dim lngTotalOtherVouchersINCL As Long
Dim lngTotalVouchersQty As Long
Dim lngTotalCNRedeemed As Long
Dim lngTotalCNIssued As Long
Dim lngTotalAccountCredits As Long
Dim lngTotalDepositsRedeemed As Long
Dim lngTotalAccountsReceivable As Long
Dim lngTotalAccountsPaid As Long
Dim lngTotalDepositsReceived As Long
Dim lngTotalDirectDeposits As Long
Dim lngTotalRefunds As Long
Dim lngTotalCreditCards As Long
Dim lngTotalCreditCardsRefunds As Long
'Dim lngTotalSalesIncl As Long
'Dim lngTotalSalesExcl As Long
Dim lngTotalDiscountIncl As Long
Dim lngTotalOtherVouchersValue As Long
Dim lngTotalLoyaltyVouchersValue As Long
Dim lngTotalLoyaltyVouchersQty As Long
Dim lngTotalDepositsrefunded As Long
Dim lngTotalDepositsRefundedCash As Long
Dim lngTotalChangeGiven As Long
Dim lngTotalCashRefunds As Long
Dim lngTotalPettyCashWithdrawal As Long
Dim lngTotalPettyCashCredit As Long
Dim lngTotalCashInDrawer As Long
Dim lngTotalPayments As Long
Dim lngWages As Long
Dim lngSickLeave As Long
Dim lngLeavePay As Long
Dim lngSales As Long
Dim lngSalesINC As Long
Dim lngSalesOnAccount As Long
Dim lngSalesAll As Long
Dim bCashUpExtended As Boolean
Dim bCashUpExtendedBB As Boolean
Dim dblCashRetained As Double
Dim dblCashReturned As Double
Dim arVoucherValue() As Long
Dim arVoucherLabel() As String

Dim rsBlindCashup As ADODB.Recordset

Dim strTotalVouchersRedeemed As String
Dim strZID As String
Dim strXID As String

Dim dteStart As Date
Dim dteEnd As Date
Dim strSupervisor As String
Dim strStation As String
Dim dtePreviousStart As Date
Dim dtePreviousEnd As Date



Public Sub component(pStartDate As Date, pENdDate As Date, pSUpervisor As String, pStation As String, pdtePreviousStart As Date, pdtePreviousEnd As Date)
    dteStart = pStartDate
    dteEnd = pENdDate
    strSupervisor = pSUpervisor
    dtePreviousStart = pdtePreviousStart
    dtePreviousEnd = pdtePreviousEnd
    strStation = pStation
End Sub
Public Sub SelectSession(pType As String, PID As String)
    If pType = "Z" Then
        strZID = PID
        strXID = ""
    ElseIf pType = "X" Then
        strXID = PID
        strZID = ""
    End If
End Sub

Public Property Get TotalCash() As Long
    TotalCash = lngTotalCash
End Property
Public Property Get TotalCashF() As String
    TotalCashF = Format(lngTotalCash / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalCashDec() As Double
    TotalCashDec = (CDbl(lngTotalCash) / oPC.Configuration.DefaultCurrency.Divisor)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalDirectDeposits() As Long
    TotalDirectDeposits = lngTotalDirectDeposits
End Property
Public Property Get TotalDirectDepositsF() As String
    TotalDirectDepositsF = Format(lngTotalDirectDeposits / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalDirectDepositsDec() As Double
    TotalDirectDepositsDec = (CDbl(lngTotalDirectDeposits) / oPC.Configuration.DefaultCurrency.Divisor)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCheques() As Long
    TotalCheques = lngTotalCheques
End Property
Public Property Get TotalChequesF() As String
    TotalChequesF = Format(lngTotalCheques / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalChequesDec() As Double
    TotalChequesDec = (CDbl(lngTotalCheques) / oPC.Configuration.DefaultCurrency.Divisor)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalVouchersRedeemed() As Long
    TotalVouchersRedeemed = lngTotalVouchersRedeemed
End Property
Public Property Get TotalVouchersRedeemedF() As String
    TotalVouchersRedeemedF = Format(lngTotalVouchersRedeemed / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalVouchersRedeemedDec() As Double
    TotalVouchersRedeemedDec = TotalVouchersRedeemedDbl
End Property
Public Property Get TotalVouchersRedeemedDecF() As String
    TotalVouchersRedeemedDecF = Format(TotalVouchersRedeemedDec, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get TotalVouchersSoldDec() As Double
    TotalVouchersSoldDec = CDbl(lngTotalVouchersINCL) / oPC.Configuration.DefaultCurrency.Divisor
End Property
'''''''''''''''''''''''
Public Property Get TotalVouchersSold() As Long
    TotalVouchersSold = lngTotalVouchersINCL
End Property
Public Property Get TotalVouchersSoldF() As String
    TotalVouchersSoldF = Format(lngTotalVouchersINCL / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get TotalOtherVouchersSold() As Long
    TotalOtherVouchersSold = lngTotalOtherVouchersINCL
End Property
Public Property Get TotalOtherVouchersSoldF() As String
    TotalOtherVouchersSoldF = Format(TotalOtherVouchersSold / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get TotalCNRedeemed() As Long
    TotalCNRedeemed = lngTotalCNRedeemed
End Property
Public Property Get TotalCNRedeemedF() As String
    TotalCNRedeemedF = Format(lngTotalCNRedeemed / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalDepositsRedeemed() As Long
    TotalDepositsRedeemed = lngTotalDepositsRedeemed
End Property
Public Property Get TotalDepositsRedeemedF() As String
    TotalDepositsRedeemedF = Format(lngTotalDepositsRedeemed / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalAccountsReceivable() As Long
    TotalAccountsReceivable = lngTotalAccountsReceivable
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalAccountsPaidF() As String
    TotalAccountsPaidF = Format(TotalAccountsPaid / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get TotalAccountsPaid() As Long
    TotalAccountsPaid = lngTotalAccountsPaid
End Property

Public Property Get TotalAccountsReceivableF() As String
    TotalAccountsReceivableF = Format(TotalAccountsReceivable / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalDepositsReceived() As Long
    TotalDepositsReceived = lngTotalDepositsReceived
End Property
Public Property Get TotalDepositsReceivedF() As String
    TotalDepositsReceivedF = Format(lngTotalDepositsReceived / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get TotalCNIssuedDec() As Double
    TotalCNIssuedDec = CDbl(lngTotalCNIssued) / oPC.Configuration.DefaultCurrency.Divisor
End Property
'''''''''''''''''''''''
Public Property Get TotalCNIssued() As Long
    TotalCNIssued = lngTotalCNIssued
End Property
Public Property Get TotalCNIssuedF() As String
    TotalCNIssuedF = Format(lngTotalCNIssued / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalAcountCredits() As Long
    TotalAcountCredits = lngTotalAccountCredits
End Property
Public Property Get TotalAccountCreditsF() As String
    TotalAccountCreditsF = Format(lngTotalAccountCredits / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalRefunds() As Long
    TotalRefunds = lngTotalRefunds
End Property
Public Property Get TotalRefundsF() As String
    TotalRefundsF = Format(lngTotalRefunds / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCreditCards() As Long
    TotalCreditCards = lngTotalCreditCards
End Property
Public Property Get TotalCreditCardsF() As String
    TotalCreditCardsF = Format(lngTotalCreditCards / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalCreditCardsDec() As Double
    TotalCreditCardsDec = (CDbl(lngTotalCreditCards) / oPC.Configuration.DefaultCurrency.Divisor)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCreditCardsRefunds() As Long
    TotalCreditCardsRefunds = lngTotalCreditCardsRefunds
End Property
Public Property Get TotalCreditCardsRefundsF() As String
    TotalCreditCardsRefundsF = Format(lngTotalCreditCardsRefunds / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalCreditCardsRefundsDec() As Double
    TotalCreditCardsRefundsDec = (CDbl(lngTotalCreditCardsRefunds) / oPC.Configuration.DefaultCurrency.Divisor)
End Property

'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCreditCardsNett() As Long
    TotalCreditCardsNett = lngTotalCreditCards + lngTotalCreditCardsRefunds   'lngTotalCreditCardsRefunds have sign
End Property
Public Property Get TotalCreditCardsNettF() As String
    TotalCreditCardsNettF = Format(TotalCreditCardsNett / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalCreditCardsNettDec() As Double
    TotalCreditCardsNettDec = (CDbl(TotalCreditCardsNett) / oPC.Configuration.DefaultCurrency.Divisor)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get SalesINC() As Long
    SalesINC = lngSalesINC
End Property
Public Property Get TotalSalesInclF() As String
    TotalSalesInclF = Format(SalesINC / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get SalesOnAccount() As Double
    SalesOnAccount = lngSalesOnAccount
End Property
Public Property Get TotalSalesOnAccountF() As String
    TotalSalesOnAccountF = Format(SalesOnAccount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get SalesAll() As Long
    SalesAll = lngSalesOnAccount + lngSalesINC
End Property
Public Property Get TotalSalesAllF() As String
    TotalSalesAllF = Format(SalesAll / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalDiscountIncl() As Long
    TotalDiscountIncl = lngTotalDiscountIncl
End Property
Public Property Get TotalDiscountInclF() As String
    TotalDiscountInclF = Format(lngTotalDiscountIncl / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
'Public Property Get TotalSalesExcl() As Long
'    TotalSalesExcl = lngTotalSalesExcl
'End Property
'Public Property Get TotalSalesExclF() As String
'    TotalSalesExclF = Format(lngTotalSalesExcl / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
''''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalLoyaltyVouchersValue() As Long
    TotalLoyaltyVouchersValue = lngTotalLoyaltyVouchersValue
End Property
Public Property Get TotalLoyaltyVouchersValueF() As String
    TotalLoyaltyVouchersValueF = Format(lngTotalLoyaltyVouchersValue / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalLoyaltyQty() As Long
    TotalLoyaltyQty = lngTotalLoyaltyVouchersQty
End Property
'''''''''''''''''''''''


'''''''''''''''''''''''
Public Property Get TotalDepositsrefunded() As Long
    TotalDepositsrefunded = lngTotalDepositsrefunded
End Property
Public Property Get TotalDepositsrefundedF() As String
    TotalDepositsrefundedF = Format(lngTotalDepositsrefunded / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalChangeGiven() As Long
    TotalChangeGiven = lngTotalChangeGiven
End Property
Public Property Get TotalChangeGivenF() As String
    TotalChangeGivenF = Format(lngTotalChangeGiven / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalChangeGivenDec() As Double
    TotalChangeGivenDec = CDbl(lngTotalChangeGiven) / oPC.Configuration.DefaultCurrency.Divisor
End Property
Public Property Get TotalChangeGivenDecF() As String
    TotalChangeGivenDecF = Format(TotalChangeGivenDec, oPC.Configuration.DefaultCurrency.FormatString)
End Property
''''''''''''''''''''''''
''''''''''''''''''''''''
'Public Property Get TotalCashInDrawer() As Long
'    TotalCashInDrawer = lngTotalCash - lngTotalChangeGiven - lngTotalCashRefunds
'End Property
'Public Property Get TotalCashInDrawerF() As String
'    TotalCashInDrawerF = Format(TotalCashInDrawer / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalPettyCashWithdrawals() As Long
    TotalPettyCashWithdrawals = lngTotalPettyCashWithdrawal
End Property
Public Property Get TotalPettyCashWithdrawalsF() As String
    TotalPettyCashWithdrawalsF = Format(TotalPettyCashWithdrawals / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalPettyCashCredits() As Long
    TotalPettyCashCredits = lngTotalPettyCashCredit
End Property
Public Property Get TotalPettyCashCreditsF() As String
    TotalPettyCashCreditsF = Format(TotalPettyCashCredits * -1 / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalPettyCashNett() As Long
    TotalPettyCashNett = lngTotalPettyCashWithdrawal + lngTotalPettyCashCredit
End Property
Public Property Get TotalPettyCashNettDec() As Double
    TotalPettyCashNettDec = CDbl(FNN(lngTotalPettyCashWithdrawal) + FNN(lngTotalPettyCashCredit)) / oPC.Configuration.DefaultCurrency.Divisor
End Property
Public Property Get TotalPettyCashNettF() As String
    TotalPettyCashNettF = Format(TotalPettyCashNett / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCashInDrawer() As Long
    TotalCashInDrawer = lngTotalCashInDrawer
End Property
Public Property Get TotalCashInDrawerF() As String
    TotalCashInDrawerF = Format(TotalCashInDrawer / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get TotalCashInDrawerDec() As Double
    TotalCashInDrawerDec = CDbl(lngTotalCashInDrawer) / oPC.Configuration.DefaultCurrency.Divisor
End Property
Public Property Get TotalCashInDrawerDecF() As String
    TotalCashInDrawerDecF = Format(TotalCashInDrawerDec, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalPayments() As Long
    TotalPayments = lngTotalPayments
End Property
Public Property Get TotalPaymentsF() As String
    TotalPaymentsF = Format(TotalPayments / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get TotalCOGSDec() As Double
    TotalCOGSDec = CDbl(lngCOGS) / oPC.Configuration.DefaultCurrency.Divisor
End Property

'''''''''''''''''''''''
Public Property Get TotalCOGS() As Long
    TotalCOGS = lngCOGS
End Property
Public Property Get TotalCOGSF() As String
    TotalCOGSF = Format(TotalCOGS / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get TotalWagesDec() As Double
    TotalWagesDec = CDbl(lngWages) / oPC.Configuration.DefaultCurrency.Divisor
End Property
Public Property Get TotalWages() As Long
    TotalWages = lngWages
End Property
Public Property Get TotalWagesF() As String
    TotalWagesF = Format(TotalWages / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get TotalSickLeaveDec() As Double
    TotalSickLeaveDec = CDbl(lngSickLeave) / oPC.Configuration.DefaultCurrency.Divisor
End Property
Public Property Get TotalSickLeave() As Long
    TotalSickLeave = lngSickLeave
End Property
Public Property Get TotalSickLeaveF() As String
    TotalSickLeaveF = Format(TotalSickLeave / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get TotalLeavePayDec() As Double
    TotalLeavePayDec = CDbl(lngLeavePay) / oPC.Configuration.DefaultCurrency.Divisor
End Property
Public Property Get TotalLeavePay() As Long
    TotalLeavePay = lngLeavePay
End Property
Public Property Get TotalLeavePayF() As String
    TotalLeavePayF = Format(TotalLeavePay / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
Public Property Get SalesDec() As Double
    SalesDec = CDbl(lngSalesINC) / oPC.Configuration.DefaultCurrency.Divisor
End Property
'''''''''''''''''''''''
Public Property Get Sales() As Long
    Sales = lngSales
End Property
Public Property Get SalesF() As String
    SalesF = Format(Sales / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get CashF() As String
    CashF = Format((Sales + (dblCashReturned * 100) - (dblCashRetained * 100)) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

'''''''''''''''''''''''
''''''''''''''''''''''' + (dblCashReturned * 100) - (dblCashRetained * 100)

Public Function Calculate()
    On Error GoTo errHandler
Dim OpenResult As Integer
Dim lngCount As Long
Dim rs As ADODB.Recordset
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    lngCOGS = FetchTotalCostOfGoodsSold
    lngTotalCash = FetchTotalCash
    lngTotalCheques = FetchTotalCheques
    lngTotalCNRedeemed = FetchTotalCNRedeemed
    lngTotalCNIssued = FetchTotalCNIssued
    
    lngTotalDepositsRedeemed = FetchTotalDepositsRedeemed
    lngTotalDepositsReceived = FetchTotalDepositsAccepted
    
    lngTotalRefunds = FetchTotalRefunds
    
    lngTotalCreditCards = FetchTotalCreditCards
    lngTotalCreditCardsRefunds = FetchTotalCreditCardsRefunds   'this figure includes refunds of deposits
    
    lngTotalCashRefunds = FetchTotalCashRefunds
    
    lngTotalDirectDeposits = FetchDirectDeposits
    
    lngTotalPettyCashWithdrawal = FetchTotalPettyCashWithdrawal
    lngTotalPettyCashCredit = FetchTotalPettyCashCredit
    
    lngTotalAccountsReceivable = FetchTotalAccountsReceivable
    lngTotalAccountsPaid = FetchTotalAccountsPaid
    lngTotalAccountCredits = FetchTotalAccountCredits  ' Removed a ' * 100 ' factor that was yielding too great a result on credit notes
    lngSalesOnAccount = FetchTotalSalesOnAccount
   ' FetchTotalSales
    FetchTotalDiscount
    FetchTotalLoyaltyVouchers
    FetchTotalDepositsRefunded
 '   FetchTotalDepositsRefundedCash
    FetchTotalChangeGiven
    FetchTotalVouchersSold
    
    FetchTotalCashRetainedandReturned
    
    Set rs = FetchTotalVouchersRedeemed(oPC.VouchersReportTogether, lngCount)
    lngTotalVouchersRedeemed = 0
    Do While Not rs.eof
        i = i + 1
        lngTotalVouchersRedeemed = lngTotalVouchersRedeemed + FNN(rs.Fields(0))
        rs.MoveNext
    Loop
    rs.Close
    
'lngTotalCash is all cash received as a result of sales. We need to produce a
'figure representing cash in till, so we must adjust by subtracting cash withdrawn as petty cash,
'adding cash returned as petty cash credits, adding cash received as deposits, subtracting cash returned as deposits refunded
'and Subtracting cash returns
    lngTotalCashInDrawer = lngTotalCash - lngTotalChangeGiven - lngTotalPettyCashWithdrawal + ((lngTotalPettyCashCredit) * -1) '+ lngTotalCashRefunds     '+ lngTotalDepositsReceived - lngTotalDepositsrefunded
    lngSales = lngTotalCashInDrawer + lngTotalCheques + lngTotalCreditCards + lngTotalCreditCardsRefunds  'note deposit refunds are not part of this as a deposit is not part of a sale in the first place
    lngTotalPayments = lngSales + lngTotalDirectDeposits
    
    lngSalesINC = lngTotalPayments - lngTotalVouchersINCL + lngTotalVouchersRedeemed - lngTotalAccountCredits _
                    + lngTotalAccountsReceivable + lngTotalDepositsRedeemed - lngTotalDepositsReceived - lngTotalAccountsPaid _
                    + lngTotalDepositsrefunded + (lngTotalPettyCashWithdrawal _
                    + (lngTotalPettyCashCredit)) - lngTotalCNIssued
   '- lngTotalDepositsReceived
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.Calculate"
End Function
Public Function TotalVouchersRedeemedEx() As Long
Dim lngTmp As Long
Dim OpenResult As Integer
Dim rs As ADODB.Recordset
Dim rsPettyCash As ADODB.Recordset

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = FetchTotalVouchersRedeemed()
    lngTmp = 0
    Do While Not rs.eof
        i = i + 1
        ReDim Preserve arVoucherValue(i)
        arVoucherValue(i) = FNN(rs.Fields(0))
        lngTmp = lngTmp + arVoucherValue(i)
        rs.MoveNext
    Loop
    rs.Close
    TotalVouchersRedeemedEx = lngTmp
End Function
Public Function TotalVouchersRedeemedDbl() As Double
    TotalVouchersRedeemedDbl = CDbl(TotalVouchersRedeemedEx) / 100
End Function

Public Function PrintCashup(pType As String, Optional pProvisional As Boolean, Optional ZID As String)
    On Error GoTo errHandler
Dim ar As Object
Dim oSQL As New z_SQL
Dim rs As ADODB.Recordset
Dim rsPettyCash As ADODB.Recordset
Dim i As Integer
Dim strLabel As String
Dim strValues As String
Dim lngCount As Long
Dim OpenResult As Integer
Dim iPos As Integer
Dim mXID As String
Dim ActualCashTaken As Double

    If oPC.GetProperty("Cashup_Extended") = "TRUE" Then
        If oPC.GetProperty("DeliveryStyle") = "BB" Then
            Set ar = New arCashupExt2
            bCashUpExtended = True
            bCashUpExtendedBB = True
            If strXID > "" Then
                mXID = strXID
                Set rsBlindCashup = New ADODB.Recordset
                rsBlindCashup.CursorLocation = adUseClient
                oSQL.LoadBlindCashup rsBlindCashup, mXID
            End If
            
        Else
            Set ar = New arCashupExt
            bCashUpExtended = True
        End If
    Else
        Set ar = New arCashup
        bCashUpExtended = False
        bCashUpExtendedBB = True
    End If
    
    ar.Printer.Orientation = ddOPortrait
    ar.Width = 12000
    ar.Height = 8000
    If pType = "Z" Then
        ar.Caption = "Daily Cashup  " & "session started : " & Format(dteStart, "dd/mm/yyyy HH:NN") & "     " & "Ended : " & Format(dteEnd, "dd/mm/yyyy HH:NN")
    Else
        If pProvisional Then
            ar.Caption = "Operator Cashup  " & "session started : " & Format(dteStart, "dd/mm/yyyy HH:NN") & "     " & "Ended : PROVISIONAL OPERATOR CASHUP"
        Else
            ar.Caption = "Operator Cashup  " & "session started : " & Format(dteStart, "dd/mm/yyyy HH:NN") & "     " & "Ended : " & Format(dteEnd, "dd/mm/yyyy HH:NN")
        End If
    End If
    ar.lblHeading.Caption = ar.Caption
    
    ar.lblStation.Caption = "Station : " & strStation
    ar.lblSupervisor.Caption = "Supervisor: " & strSupervisor
    ar.lblPrinted.Caption = "Printed: " & Format(Now(), "dd/mm/yyyy HH:NN")
    ar.fCash.text = TotalCashInDrawerF
    If oPC.GetProperty("DeliveryStyle") = "BB" And oPC.GetProperty("Cashup_Extended") = "TRUE" And strXID > "" Then
        ActualCashTaken = FNDBL(rsBlindCashup.Fields("TotalCash")) - FNDBL(rsBlindCashup.Fields("FloatValueatStart")) + FNDBL(rsBlindCashup.Fields("Float"))
        ar.fCashA.text = Format(FNDBL(ActualCashTaken), "###,##0.00")
    End If
    
    ar.fCheques.text = Me.TotalChequesF
    If oPC.GetProperty("DeliveryStyle") = "BB" And oPC.GetProperty("Cashup_Extended") = "TRUE" And strXID > "" Then
        ar.fChequesA.text = Format(FNDBL(rsBlindCashup.Fields("Cheques")), "###,##0.00")
    End If
    
    ar.fDailySalesIncl.text = CashF  'SalesF
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = FetchTotalVouchersRedeemed(oPC.VouchersReportTogether, lngCount)
    
    lngTotalVouchersRedeemed = 0
    Do While Not rs.eof
        i = i + 1
        ReDim Preserve arVoucherValue(i)
        ReDim Preserve arVoucherLabel(i)
        arVoucherValue(i) = FNN(rs.Fields(0))
        lngTotalVouchersRedeemed = lngTotalVouchersRedeemed + arVoucherValue(i)
        arVoucherLabel(i) = FNS(rs.Fields(1))
        strLabel = strLabel & IIf(Len(strLabel) > 0, vbCrLf, "") & arVoucherLabel(i)
        strValues = strValues & IIf(Len(strValues) > 0, vbCrLf, "") & Format(arVoucherValue(i) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        rs.MoveNext
    Loop
    rs.Close

    lngSalesINC = lngTotalPayments - lngTotalVouchersINCL + lngTotalVouchersRedeemed - lngTotalAccountCredits _
                    + lngTotalAccountsReceivable + lngTotalDepositsRedeemed - lngTotalDepositsReceived - lngTotalAccountsPaid _
                    + lngTotalDepositsrefunded + (lngTotalPettyCashWithdrawal _
                    + (lngTotalPettyCashCredit)) - lngTotalCNIssued
                    
                    '- lngTotalDepositsReceived
    lngSalesAll = lngSalesINC + lngSalesOnAccount
    Set rsPettyCash = FetchPettyCash
    ar.component rsPettyCash, TotalPettyCashCreditsF, TotalPettyCashWithdrawalsF, TotalPettyCashNettF, strZID
    Set rsPettyCash = FetchPettyCash
   
    lngWages = 0
    lngSickLeave = 0
    lngLeavePay = 0
    If rsPettyCash.eof And rsPettyCash.BOF Then
    Else
    Do While Not rsPettyCash.eof
        iPos = InStr(1, FNS(rsPettyCash.Fields(1)), ":")
        Select Case Left(FNS(rsPettyCash.Fields(1)), iPos - 1)
        Case "W"
            lngWages = lngWages + FNN(rsPettyCash.Fields(2))
        Case "SL"
            lngWages = lngSickLeave + FNN(rsPettyCash.Fields(2))
        Case "LP"
            lngWages = lngLeavePay + FNN(rsPettyCash.Fields(2))
        End Select
        rsPettyCash.MoveNext
    Loop
    End If
    
    
    
    ar.fVouchers.text = strValues
    ar.lblVoucher.text = strLabel
    If oPC.GetProperty("DeliveryStyle") = "BB" And oPC.GetProperty("Cashup_Extended") = "TRUE" And strXID > "" Then
        ar.fVouchersRedeemedA.text = Format(FNDBL(rsBlindCashup.Fields("Vouchers")), "###,##0.00")
    End If
    ar.fWWVouchersSold.text = Me.TotalVouchersSoldF
    ar.fCNIssued.text = TotalCNIssuedF
   ' ar.fCNRedeemed = TotalCNRedeemedF
    ar.fDepReceived.text = TotalDepositsReceivedF
    ar.fDepRedeemed.text = TotalDepositsRedeemedF
    ar.fACRec.text = TotalAccountsReceivableF
    ar.fAccountReceipts.text = TotalAccountsPaidF
    ar.fAccountCreditnotes.text = TotalAccountCreditsF
    ar.fCreditCards.text = TotalCreditCardsF
    ar.fSalesIncl.text = TotalSalesInclF
    If bCashUpExtended = True Then
        ar.fSalesOnAccount.text = TotalSalesOnAccountF
        ar.fSalesAll.text = TotalSalesAllF
        If bCashUpExtendedBB Then   'ONly for BB so far
            ar.fRetained.text = Format(dblCashRetained, "###,##0.00")
            ar.fReturned.text = Format(dblCashReturned, "###,##0.00")
        End If
    End If
    ar.fCreditCardsRefunded.text = TotalCreditCardsRefundsF
    ar.fCreditCardsNett.text = TotalCreditCardsNettF
    If oPC.GetProperty("DeliveryStyle") = "BB" And oPC.GetProperty("Cashup_Extended") = "TRUE" And strXID > "" Then
        ar.fCreditCardsNettA.text = Format(FNDBL(rsBlindCashup.Fields("CreditCards")) + FNDBL(rsBlindCashup.Fields("DebitCards")), "###,##0.00")
    End If
    ar.fTotalDiscount.text = TotalDiscountInclF
    ar.fLVouchersIssued.text = TotalLoyaltyVouchersValueF
    ar.fLoyaltyVouchersQty.text = "(" & TotalLoyaltyQty & ")"
    ar.fDepositsRefunded.text = TotalDepositsrefundedF
    ar.fDirectDeposits.text = TotalDirectDepositsF
    If oPC.GetProperty("DeliveryStyle") = "BB" And oPC.GetProperty("Cashup_Extended") = "TRUE" And strXID > "" Then
        ar.fDirectDepositsA.text = Format(FNDBL(rsBlindCashup.Fields("DirectDeposits")), "###,##0.00")
    End If
    ar.fTotalPayments.text = Format((TotalPayments - (dblCashRetained * 100) + (dblCashReturned * 100)) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString) 'TotalPaymentsF
    ar.fCOGS.text = TotalCOGSF
    If CLng(dtePreviousStart) <> 0 Then
        If pType = "Z" Then
            ar.lblFoot.Caption = "Previous session started : " & Format(dtePreviousStart, "dd/mm/yyyy HH:NN") & "     " & "Ended : " & Format(dtePreviousEnd, "dd/mm/yyyy HH:NN") & "   Supervisor: " & strSupervisor
        Else
            ar.lblFoot.Caption = "Previous session started : " & Format(dtePreviousStart, "dd/mm/yyyy HH:NN") & "     " & "Ended : " & Format(dtePreviousEnd, "dd/mm/yyyy HH:NN") & "   Supervisor: " & strSupervisor
        End If
    End If
    
    ar.Show vbModal
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.PrintCashup(pType)", pType, , , "line number", Array(Erl())
End Function

Public Function FetchTotalSalesOnAccount() As Long
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalSalesOnAccount"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@StartDate", adDate, adParamInput, , dteStart)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@EndDate", adDate, adParamInput, , dteEnd)  'This is not a mistake (see the SP)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adDouble, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalSalesOnAccount = FNDBL(cmd.Parameters(2)) * oPC.Configuration.DefaultCurrency.Divisor
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalSalesOnAccount"
End Function
Public Function FetchTotalCostOfGoodsSold()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCOGS"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalCostOfGoodsSold = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCostOfGoodsSold"
End Function

Public Function FetchTotalCash()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCash"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalCash = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCash"
End Function
Public Sub FetchTotalCashRetainedandReturned()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCashRetainedAndReturned"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Retained", adInteger, adParamOutput)
            par.NumericScale = 2
            par.Precision = 12
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Returned", adInteger, adParamOutput)
            par.NumericScale = 2
            par.Precision = 12
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    dblCashRetained = CDbl(cmd.Parameters(2))
    dblCashReturned = CDbl(cmd.Parameters(3))
    Set cmd = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCashRetainedandReturned"
End Sub
Public Property Get CashRetained() As Double
    CashRetained = dblCashRetained
End Property
Public Property Get CashReturned() As Double
    CashReturned = dblCashReturned
End Property
Public Function FetchTotalCheques()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCheques"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalCheques = FNN(cmd.Parameters(2))
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCheques"
End Function

Public Function FetchTotalVouchersRedeemed(Optional strTypes As String, Optional pResult As Long) As ADODB.Recordset
    On Error GoTo errHandler
Dim lngResult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset
    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalVouchers"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VoucherTypes", adVarChar, adParamInput, 40, "") 'strTypes)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    Set rs = cmd.execute
    Set cmd = Nothing
    Set par = Nothing
    Set FetchTotalVouchersRedeemed = rs
    Set rs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalVouchersRedeemed"
End Function
Public Function FetchPettyCash() As ADODB.Recordset
    On Error GoTo errHandler
Dim lngResult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset
    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetPettyCash"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    Set rs = cmd.execute
    Set cmd = Nothing
    Set par = Nothing
    Set FetchPettyCash = rs
    Set rs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchPettyCash"
End Function

Public Function FetchTotalCNRedeemed()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCNRedeemed"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalCNRedeemed = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCNRedeemed"
End Function

Public Function FetchTotalCNIssued()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCNIssued"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalCNIssued = cmd.Parameters(2) * -1  'Keep this number positive for consistancy
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCNIssued"
End Function
Public Function FetchTotalAccountCredits()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalAccountCredits"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adNumeric, adParamOutput)
            par.NumericScale = 4
            par.Precision = 15
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalAccountCredits = cmd.Parameters(2) '* -1  'Keep this number positive for consistancy
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalAccountCredits"
End Function
Public Function FetchTotalDepositsAccepted()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalDepositsAccepted"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalDepositsAccepted = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalDepositsAccepted"
End Function

Public Function FetchTotalDepositsRedeemed()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalDepositsRedeemed"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalDepositsRedeemed = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalDepositsRedeemed"
End Function

Public Function FetchTotalRefunds()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalRefunds"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalRefunds = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalRefunds"
End Function
Public Function FetchTotalCreditCards()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCreditCards"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalCreditCards = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCreditCards"
End Function
Public Function FetchTotalCreditCardsRefunds()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCreditCardsRefunds"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalCreditCardsRefunds = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCreditCardsRefunds"
End Function
Public Function FetchTotalCashRefunds()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalCashRefunds"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Val", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalCashRefunds = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalCashRefunds"
End Function

'Public Function FetchTotalSales()
'    On Error GoTo ErrHandler
'Dim iresult As Long
'Dim cmd As ADODB.Command
'Dim par As ADODB.Parameter
'
'    Set cmd = New ADODB.Command
'    cmd.CommandText = "spcu_GetTotalSales"
'    cmd.CommandType = adCmdStoredProc
'
'    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@ValIncl", adInteger, adParamOutput)
'    cmd.Parameters.Append par
'    Set par = cmd.CreateParameter("@ValExcl", adInteger, adParamOutput)
'    cmd.Parameters.Append par
'
'    cmd.ActiveConnection = oPC.COShort
'    cmd.execute
'
'    lngTotalSalesIncl = cmd.Parameters(2)
'    lngTotalSalesExcl = cmd.Parameters(3)
'    Set cmd = Nothing
'
'    Exit Function
'ErrHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_Cashup.FetchTotalSales"
'End Function
Public Function FetchTotalVouchersSold()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalVouchersSold"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ValIncl", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ValExcl", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Qty", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    lngTotalVouchersINCL = cmd.Parameters(2)
    lngTotalVouchersEXCL = cmd.Parameters(3)
    lngTotalVouchersQty = cmd.Parameters(4)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalSales"
End Function

Public Function FetchTotalDiscount()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalDiscount"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ValIncl", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    lngTotalDiscountIncl = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalDiscount"
End Function

Public Function FetchTotalLoyaltyVouchers()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalLoyaltyVouchers"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TotalLoyalty", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Qty", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    lngTotalLoyaltyVouchersValue = cmd.Parameters(2)
    lngTotalLoyaltyVouchersQty = cmd.Parameters(3)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalLoyaltyVouchers"
End Function

Public Function FetchTotalDepositsRefunded()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalDepositsRefunded"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Total", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    lngTotalDepositsrefunded = cmd.Parameters(2) * -1   'To keep consistency
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalDepositsRefunded"
End Function
Public Function FetchTotalDepositsRefundedCash()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalDepositsRefundedCash"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Total", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    lngTotalDepositsRefundedCash = cmd.Parameters(2) * -1   'To keep consistency
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalDepositsRefunded"
End Function
Public Function FetchTotalChangeGiven()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalChange"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Total", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    lngTotalChangeGiven = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalChangeGiven"
End Function

Public Function FetchTotalPettyCashWithdrawal()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalPettyCashWithdrawals"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TotalPC", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalPettyCashWithdrawal = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalPettyCashWithdrawal"
End Function
Public Function FetchTotalPettyCashCredit()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalPettyCashCredits"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TotalPC", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalPettyCashCredit = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.spcu_GetTotalPettyCashCredit"
End Function

Public Function FetchTotalAccountsPaid()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalAccountsPaid"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Total", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalAccountsPaid = FNN(cmd.Parameters(2))
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalAccountsReceivable"
End Function


Public Function FetchTotalAccountsReceivable()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalAccountsReceivable"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Total", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalAccountsReceivable = FNN(cmd.Parameters(2))
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalAccountsReceivable"
End Function


Public Function PrintSectionSales(pDate As Date, pZID As String)
Dim rs As ADODB.Recordset
Dim oRep As New arDailySection
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    If pZID > "" Then
        rs.Open "Select * FROM vSalesPerSectionByZID WHERE ZID = '" & pZID & "' ORDER BY SectionName ", oPC.COShort
    Else
        rs.Open "Select * FROM vSalesPerSection WHERE SaleDate = '" & ReverseDate(pDate) & "' ORDER BY SectionName ", oPC.COShort
    End If
    oRep.component rs, "Daily sales by section"
    oRep.Show vbModal
    Set oRep = Nothing
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function
Public Function PrintPTSales(pDate As Date, pZID As String)
Dim rs As ADODB.Recordset
Dim oRep As New arDailyProductType
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    If pZID > "" Then
        rs.Open "Select * FROM vSalesPerProductTypeByZID WHERE ZID = '" & pZID & "' ORDER BY ProductType ", oPC.COShort
    Else
        rs.Open "Select * FROM vSalesPerProductType WHERE SaleDate = '" & ReverseDate(pDate) & "' ORDER BY ProductType ", oPC.COShort
    End If
    oRep.component rs, "Daily sales by product type"
    oRep.Show vbModal
    Set oRep = Nothing
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function
Public Function PrintPayments(pFrom As Date, pTo As Date, pZID As String)
Dim rs As ADODB.Recordset
Dim oRep As New arExchangePayments
Dim OpenResult As Integer
Dim strHeading As String
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetSalesPayments"
    cmd.commandType = adCmdStoredProc

    Set par = cmd.CreateParameter("@FROM", adDate, adParamInput, , pFrom)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adDate, adParamInput, , pTo)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ZID", adVarChar, adParamInput, 40, CStr(pZID))
    cmd.Parameters.Append par

    cmd.ActiveConnection = oPC.COShort
    Set rs = cmd.execute

    Set cmd = Nothing

    If pZID > "" Then
        strHeading = "Payments for day-session started " & Format(pFrom, "dd-mm-yyyy HH:NN AMPM")
    Else
        If pTo > CDate("2000-01-01") Then
            strHeading = "Payments taken between " & Format(pFrom, "dd-mm-yyyy") & " and " & Format(pTo, "dd-mm-yyyy")
        Else
            strHeading = "Payments taken on " & Format(pFrom, "dd-mm-yyyy")
        End If
    End If

    oRep.component rs, strHeading
    oRep.Printer.Orientation = ddOLandscape

    oRep.Show vbModal
    Set oRep = Nothing
    rs.Close
    Set rs = Nothing

'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function

Public Function PrintDailySales(pDate As Date, pZID As String)
Dim rs As ADODB.Recordset
Dim oRep As New arDailySales
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set rs = New ADODB.Recordset
    If pZID > "" Then
        rs.Open "Select * FROM vSalesDetailsByZID WHERE ZID = '" & pZID & "' ORDER BY EXCHNUMBER ", oPC.COShort
    Else
        rs.Open "Select * FROM vSalesDetails WHERE SaleDate = '" & ReverseDate(pDate) & "' ORDER BY EXCHNUMBER ", oPC.COShort
    End If
    oRep.component rs, "Daily sales by item"
    oRep.Show vbModal
    Set oRep = Nothing
    rs.Close
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function

Public Function FetchDirectDeposits()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalDirectDeposits"
    cmd.commandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@pZID", adVarChar, adParamInput, 40, Trim(strZID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@pXID", adVarChar, adParamInput, 40, Trim(strXID))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Total", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchDirectDeposits = FNN(cmd.Parameters(2))
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchDirectDeposits"
End Function

