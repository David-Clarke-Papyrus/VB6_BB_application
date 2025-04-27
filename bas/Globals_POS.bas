Attribute VB_Name = "Globals_POS"
Option Explicit

Global Const ZACTION_LEVEL = 2
Global Const SI_TITLE = 1
Global Const SI_AUTHOR = 2
Global Const SI_UNITPR = 3
Global Const SI_QTY = 4
Global Const SI_DISC = 5
Global Const SI_PRICE = 6
Global Const SI_PID = 7
Global SQL As String
Enum eState
    eStart = 0
    eSale = 1
    eTitle = 2
    eQty = 3
    eDiscount = 4
    ePrice = 5
    elogin = 6
    ePaymentAmt = 7
    eConfirmation = 8
    eSearchCustomer = 9
    
    eXTerminate = 20
    eZTerminate = 21
    eRebuildIndexes = 22
    eHelp = 23
    ecancelsale = 24
    
    eCashRefund = 25
    ePriceCashRefund = 26
    eQtyCashRefund = 27
    eDiscountCashRefund = 28
    eConfirmationCashrefund = 29
    
    eVoid = 30
    eReviewExchanges = 31
    eShowExchange = 32
    eOPenDrawer = 33
    eStatus = 34
    eNull = 35
    ePrevious = 36
    eDelete = 37
    eDeletePayment = 38
    eShowvoucherType = 39
    eOperatorsReport = 40
    
    eCreditNote = 41
    ePriceCreditNote = 42
    eDiscountCreditNote = 43
    eQtyCreditNote = 44
    
    eRefundDeposit = 45
    eConfirmationRefundDeposit = 46
    eSearchCustomerfordepositRefund = 47
    
    eSearchCustomerforAppro = 50
    eAppro = 51
    ePriceAppro = 52
    eDiscountAppro = 53
    eQtyAppro = 54
  '  eConfirmationAppro = 55
    
    eApproReturn = 56
    eSearchCustomerforApproReturn = 57
    
    ePettyCash = 58
    ePettyCashAmt = 59
    ePettyCashConfirmation = 60
    ePettyCashReason = 61
    ePettyCashCredit = 62
    ePettyCashCreditConfirmation = 63
    ePettyCashCreditAmt = 64
    
    
    eSearchCustomerfordeposit = 65
    eDiscountDeposit = 66
    eSelectDepositLineRef = 67
    eSelectDepositLine = 68
    eSelectDepositLineForRefund = 69
    ePriceDeposit = 70
    eQtyDeposit = 71
    
    eInvoice = 72
    eInvoiceno = 73
    eInvoiceMode = 74
    eConfirmationInvoiceCollection = 75
    eConfirmationDepositRefund = 76
    eConfirmationDeposit = 77
    eConfirmationCreditNote = 78
    eEditPrice = 105
    eEditQty = 106
'Here all nodes (states) below eCollect are numbered higher for eacy reverse navigation (see SetPresentState)
    eCollect = 89
    ePaymentType_Cash = 90
    ePaymentType_Cheque = 91
    ePaymentType_CreditCard = 92
    ePaymentType_CreditVoucher = 93
    ePaymentType_CreditVoucherRef = 94
    ePaymentType_voucher = 95
    ePaymentType_ChequeRef = 96
    ePaymentType_CreditCardRef = 97
    ePaymentType_voucherRef = 98
    ePaymentType_RedeemDeposit = 99
    eRefundType_Cash = 48
    eRefundType_CreditCard = 49
    eRefundType_CreditVoucher = 100
    eEND = 101
    ePaymentType_Account = 102
    eAccountPayment = 103
    ePaymentType_DirectDeposit = 104
    eCollectRep = 105
End Enum

