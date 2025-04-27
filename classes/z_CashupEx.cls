VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_CashupEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim lngACRec As Long
Dim lngTotalCash As Long
Dim lngCOGS As Long
Dim lngTotalCheques As Long
Dim lngTotalVouchersRedeemed As Long
Dim lngTotalVouchersEXCL As Long
Dim lngTotalVouchersINCL As Long
Dim lngTotalVouchersQty As Long
Dim lngTotalCNRedeemed As Long
Dim lngTotalCNIssued As Long
Dim lngTotalAccountCredits As Long
Dim lngTotalPayments As Long
Dim lngSalesOnAccount As Long

Dim lngTotalDepositsRedeemed As Long
Dim lngTotalAccountsReceivable As Long
Dim lngTotalAccountsPaid As Long
Dim lngTotalDepositsReceived As Long
Dim lngTotalRefunds As Long
Dim lngTotalCreditCards As Long
Dim lngTotalCreditCardsRefunds As Long
Dim lngTotalSalesIncl As Long
Dim lngTotalSalesExcl As Long
Dim lngTotalDiscountIncl As Long
Dim lngTotalLoyaltyVouchersValue As Long
Dim lngTotalLoyaltyVouchersQty As Long
Dim lngTotalDepositsrefunded As Long
Dim lngTotalChangeGiven As Long
Dim lngTotalCashRefunds As Long
Dim lngTotalPettyCashWithdrawal As Long
Dim lngTotalPettyCashCredit As Long
Dim lngTotalCashInDrawer As Long
Dim lngTotalDirectDeposits As Long
Dim lngSales As Long
Dim lngSalesINC As Long
Dim lngSalesAll As Long

Dim strTotalVouchersRedeemed As String
Dim strStart As String
Dim strEnd As String

Dim dteStart As Date
Dim dteEnd As Date
Dim strSupervisor As String
Dim strStation As String
Dim dtePreviousStart As Date
Dim dtePreviousEnd As Date

Dim rs As ADODB.Recordset

Public Sub Component(dteStart As Date, dteEnd As Date)
    dteStart = dteStart
    dteEnd = dteEnd
    strStart = Year(dteStart) & "-" & Month(dteStart) & "-" & Day(dteStart)
    strEnd = Year(dteEnd) & "-" & Month(dteEnd) & "-" & Day(dteEnd)

End Sub

'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalDirectDeposits() As Long
    TotalDirectDeposits = lngTotalDirectDeposits
End Property
Public Property Get TotalDirectDepositsF() As String
    TotalDirectDepositsF = Format(lngTotalDirectDeposits / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''

Public Property Get TotalCash() As Long
    TotalCash = lngTotalCash
End Property
Public Property Get TotalCashF() As String
    TotalCashF = Format(lngTotalCash / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCheques() As Long
    TotalCheques = lngTotalCheques
End Property
Public Property Get TotalChequesF() As String
    TotalChequesF = Format(lngTotalCheques / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalVouchersRedeemed() As Long
    TotalVouchersRedeemed = lngTotalVouchersRedeemed
End Property
Public Property Get TotalVouchersRedeemedF() As String
    TotalVouchersRedeemedF = Format(lngTotalVouchersRedeemed / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalVouchersSold() As Long
    TotalVouchersSold = lngTotalVouchersINCL
End Property
Public Property Get TotalVouchersSoldF() As String
    TotalVouchersSoldF = Format(lngTotalVouchersINCL / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
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
Public Property Get TotalAccountsReceivableF() As String
    TotalAccountsReceivableF = Format(TotalAccountsReceivable / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
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
'''''''''''''''''''''''
Public Property Get TotalDepositsReceived() As Long
    TotalDepositsReceived = lngTotalDepositsReceived
End Property
Public Property Get TotalDepositsReceivedF() As String
    TotalDepositsReceivedF = Format(lngTotalDepositsReceived / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCNIssued() As Long
    TotalCNIssued = lngTotalCNIssued
End Property
Public Property Get TotalCNIssuedF() As String
    TotalCNIssuedF = Format(lngTotalCNIssued / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
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
Public Property Get TotalCOGS() As Long
    TotalCOGS = lngCOGS
End Property
Public Property Get TotalCOGSF() As String
    TotalCOGSF = Format(TotalCOGS / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
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
Public Property Get TotalCreditCards() As Long
    TotalCreditCards = lngTotalCreditCards
End Property
Public Property Get TotalCreditCardsF() As String
    TotalCreditCardsF = Format(lngTotalCreditCards / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCreditCardsRefunds() As Long
    TotalCreditCardsRefunds = lngTotalCreditCardsRefunds
End Property
Public Property Get TotalCreditCardsRefundsF() As String
    TotalCreditCardsRefundsF = Format(lngTotalCreditCardsRefunds / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
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
Public Property Get TotalAccountsPaidF() As String
    TotalAccountsPaidF = Format(TotalAccountsPaid / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Get TotalAccountsPaid() As Long
    TotalAccountsPaid = lngTotalAccountsPaid
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalCreditCardsNett() As Long
    TotalCreditCardsNett = lngTotalCreditCards + lngTotalCreditCardsRefunds   'lngTotalCreditCardsRefunds have sign
End Property
Public Property Get TotalCreditCardsNettF() As String
    TotalCreditCardsNettF = Format(TotalCreditCardsNett / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get SalesINC() As Long
    SalesINC = lngSalesINC
End Property
Public Property Get TotalSalesInclF() As String
    TotalSalesInclF = Format(SalesINC / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
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
Public Property Get TotalSalesExcl() As Long
    TotalSalesExcl = lngTotalSalesExcl
End Property
Public Property Get TotalSalesExclF() As String
    TotalSalesExclF = Format(lngTotalSalesExcl / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
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
    TotalPettyCashCreditsF = Format(TotalPettyCashCredits / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get TotalPettyCashNett() As Long
    TotalPettyCashNett = lngTotalPettyCashWithdrawal + lngTotalPettyCashCredit
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
'''''''''''''''''''''''
'''''''''''''''''''''''
Public Property Get Sales() As Long
    Sales = lngSales
End Property
Public Property Get SalesF() As String
    SalesF = Format(Sales / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
'''''''''''''''''''''''
'''''''''''''''''''''''

Public Function Calculate() As ADODB.Recordset
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
    lngTotalCreditCardsRefunds = FetchTotalCreditCardsRefunds
    
    lngTotalCashRefunds = FetchTotalCashRefunds
    
    lngTotalDirectDeposits = FetchDirectDeposits
    
    lngTotalPettyCashWithdrawal = FetchTotalPettyCashWithdrawal
    lngTotalPettyCashCredit = FetchTotalPettyCashCredit
    
    lngTotalAccountsReceivable = FetchTotalAccountsReceivable
    lngTotalAccountsPaid = FetchTotalAccountsPaid
    lngTotalAccountCredits = FetchTotalAccountCredits
    lngSalesOnAccount = FetchTotalSalesOnAccount

 '   FetchTotalSales
    FetchTotalDiscount
    FetchTotalLoyaltyVouchers
    FetchTotalDepositsRefunded
    FetchTotalChangeGiven
    FetchTotalVouchersSold
'lngTotalCash is all cash received as a result of sales. We need to produce a
'figure representing cash in till, so we must adjust by subtracting cash withdrawn as petty cash,
'adding cash returned as petty cash credits, adding cash received as deposits, subtracting cash returned as deposits refunded
'and Subtracting cash returns
    lngTotalCashInDrawer = lngTotalCash - lngTotalChangeGiven - lngTotalPettyCashWithdrawal + (lngTotalPettyCashCredit * -1) '+ lngTotalCashRefunds     '+ lngTotalDepositsReceived - lngTotalDepositsrefunded
    lngSales = lngTotalCashInDrawer + lngTotalCheques + lngTotalCreditCards + lngTotalCreditCardsRefunds
    lngTotalPayments = lngSales + lngTotalDirectDeposits
    
    
'    lngTotalCashInDrawer = lngTotalCash - lngTotalChangeGiven - lngTotalPettyCashWithdrawal + ((lngTotalPettyCashCredit) * -1) '+ lngTotalCashRefunds     '+ lngTotalDepositsReceived - lngTotalDepositsrefunded
'    lngSales = lngTotalCashInDrawer + lngTotalCheques + lngTotalCreditCards + lngTotalCreditCardsRefunds
'    lngTotalPayments = lngSales + lngTotalDirectDeposits
    
    
    
    'the following row has been moved till after the redeemed vouchers have been calculated in the 'PrintCashup' function
    Set rs = FetchTotalVouchersRedeemed(oPC.VouchersReportTogether, lngCount)
    lngTotalVouchersRedeemed = 0
    Do While Not rs.EOF
        i = i + 1
        lngTotalVouchersRedeemed = lngTotalVouchersRedeemed + FNN(rs.Fields(0))
        rs.MoveNext
    Loop
    Set rs = Nothing
    'Reload to restart cursor
    Set rs = FetchTotalVouchersRedeemed(oPC.VouchersReportTogether, lngCount)
    
    lngSalesINC = lngTotalPayments - lngTotalVouchersINCL + lngTotalVouchersRedeemed - lngTotalAccountCredits _
                    + lngTotalAccountsReceivable + lngTotalDepositsRedeemed - lngTotalDepositsReceived - lngTotalAccountsPaid _
                    + lngTotalDepositsrefunded + (lngTotalPettyCashWithdrawal _
                    + (lngTotalPettyCashCredit)) - lngTotalCNIssued      '

    Set Calculate = rs
    lngSalesAll = lngSalesINC + lngSalesOnAccount
   
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_CashupEx.Calculate"
End Function

Public Function FetchTotalCostOfGoodsSold()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalCOGS"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adDate, adParamInput, , CDate(strStart)) 'Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adDate, adParamInput, , CDate(strEnd)) 'Trim(strEnd))
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
    cmd.CommandText = "spcud_GetTotalCash"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adDate, adParamInput, , CDate(strStart)) 'Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adDate, adParamInput, , CDate(strEnd)) 'Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalCash"
End Function
Public Function FetchTotalCheques()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalCheques"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalCheques"
End Function

Public Function FetchTotalVouchersRedeemed(strTypes As String, pResult As Long) As ADODB.Recordset
    On Error GoTo errHandler
Dim lngResult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset
    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalVouchers"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@VoucherTypes", adVarChar, adParamInput, 40, "")
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    Set rs = cmd.execute
    Set cmd = Nothing
    Set par = Nothing
    Set FetchTotalVouchersRedeemed = rs
    pResult = lngResult
    'rs.Close
    Set rs = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_CashupEx.FetchTotalVouchersRedeemed"
End Function
Public Function FetchPettyCash() As ADODB.Recordset
    On Error GoTo errHandler
Dim lngResult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim rs As ADODB.Recordset
    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetPettyCash"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchPettyCash"
End Function

Public Function FetchTotalCNRedeemed()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalCNRedeemed"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalCNRedeemed"
End Function

Public Function FetchTotalCNIssued()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalCNIssued"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalCNIssued"
End Function
Public Function FetchTotalDepositsAccepted()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalDepositsAccepted"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalDepositsAccepted"
End Function

Public Function FetchTotalDepositsRedeemed()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalDepositsRedeemed"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalDepositsRedeemed"
End Function

Public Function FetchTotalRefunds()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalRefunds"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalRefunds"
End Function
Public Function FetchTotalCreditCards()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalCreditCards"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalCreditCards"
End Function
Public Function FetchTotalCreditCardsRefunds()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalCreditCardsRefunds"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalCreditCardsRefunds"
End Function
Public Function FetchTotalCashRefunds()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalCashRefunds"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalCashRefunds"
End Function

Public Function FetchTotalSales()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalSales"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ValIncl", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@ValExcl", adInteger, adParamOutput)
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TotalQty", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    lngTotalSalesIncl = cmd.Parameters(2)
    lngTotalSalesExcl = cmd.Parameters(3)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_CashupEx.FetchTotalSales"
End Function
Public Function FetchTotalVouchersSold()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalVouchersSold"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalSales"
End Function

Public Function FetchTotalDiscount()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalDiscount"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TOTALVALUEINCL", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    lngTotalDiscountIncl = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_CashupEx.FetchTotalDiscount"
End Function

Public Function FetchTotalLoyaltyVouchers()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalLoyaltyVouchers"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalLoyaltyVouchers"
End Function

Public Function FetchTotalDepositsRefunded()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalDepositsRefunded"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalDepositsRefunded"
End Function

Public Function FetchTotalChangeGiven()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalChange"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalChangeGiven"
End Function

Public Function FetchTotalPettyCashWithdrawal()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalPettyCashWithdrawals"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalPettyCashWithdrawal"
End Function
Public Function FetchTotalPettyCashCredit()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalPettyCashCredits"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.spcu_GetTotalPettyCashCredit"
End Function

Public Function FetchTotalAccountsReceivable()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalAccountsReceivable"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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
    ErrorIn "z_CashupEx.FetchTotalAccountsReceivable"
End Function
'
Public Function FetchTotalAccountCredits()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalAccountCredits"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@Total", adInteger, adParamOutput)
    cmd.Parameters.Append par
    
    cmd.ActiveConnection = oPC.COShort
    cmd.execute
    
    FetchTotalAccountCredits = cmd.Parameters(2)
    Set cmd = Nothing

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_Cashup.FetchTotalAccountCredits"
End Function

Public Function FetchTotalAccountsPaid()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalAccountsPaid"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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

Public Function FetchDirectDeposits()
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcud_GetTotalDirectDeposits"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@FROM", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@TO", adVarChar, adParamInput, 40, Trim(strEnd))
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

Public Function FetchTotalSalesOnAccount() As Long
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter

    Set cmd = New ADODB.Command
    cmd.CommandText = "spcu_GetTotalSalesOnAccount"
    cmd.CommandType = adCmdStoredProc
    
    Set par = cmd.CreateParameter("@StartDate", adVarChar, adParamInput, 40, Trim(strStart))
    cmd.Parameters.Append par
    Set par = cmd.CreateParameter("@EndDate", adVarChar, adParamInput, 40, Trim(strEnd))  'This is not a mistake (see the SP)
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
    ErrorIn "z_CashupEx.FetchTotalSalesOnAccount"
End Function

