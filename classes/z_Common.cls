VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_Common"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum enumDocumentKind
    enExcel = 1
    enPDF = 2
    enRTF = 3
    enTabDelimited = 4
    enCommaDelimited = 5
End Enum
Public Enum enumImportType
    enStockImport = 1
    encustomerImport = 2
    enSupplierImport = 3
    enStockCategoriesImport = 4
    enSalesOrdersImport = 5
    enPurchaseOrdersImport = 6
    enCustomerCategoriesImport = 7
    enGRNImport = 8
    enPTImport = 9
    enWSSalesImport = 10
    enClipboardImport = 11
    enBookfindFeed = 12
    enBankStatement = 13
    enSellingPricesUpdate = 14
End Enum

''







Public Enum OpTypes
    Dayend = 1
    Compact = 2
    ClearOrders = 3
    Export = 4
    SupplierChange = 5
    SimplifyTitles = 6
    ClearOldData = 7
    AllocationStockToOrders = 8
    NielsenSales = 9
    BookDataWash = 10
    LoyaltyScheme = 11
    StockSharing = 12
    PastelExportDebtors = 13
    PastelExportCreditors = 14
    PastelExportCustomers = 15
    PastelImportCustomers = 16
    
    enExportGroup = 20
End Enum


Public Enum enCodeType
    enCustomerOrder = 19
    enAppro = 13
    enApproReturn = 14
    enCashSale = 20
    enSupplierOrder = 16
    enStockTake = 17
    enOOS = 25
    enInvoice = 21
    enProduct = 23
    enCreditNote = 26
    enDelivery = 18
    enTransfer = 22
    enReturn = 28
    enDebtorsJournal = 25
    enGDN = 26
    en230 = 230
    en240 = 240
    en250 = 250
    en260 = 260
    en270 = 270
    en280 = 280
End Enum

Public Enum enumExchangeType
    eSaleType = 1
    ereturntype = 2
    ePettyCashType = 3
    eDepositType = 4
    eReturnDepositType = 5
    eApproType = 6
    eCreditVoucherType = 7
    ePettyCashCreditType = 8
    eApproReturnType = 9
    eInvoiceType = 10
    eOrdequestType = 11
    eVoidAction = 12
    eXEND = 13
    eZEND = 14
    eOpenDrawerType = 15
    eAccountPaymentType = 16
    eAccountCreditNoteType = 17
    eAccountSaleType = 18
End Enum

Public Enum enStatus
    stCOMPLETE = 4
    stInProcess = 2
    stVOID = 1
    stPROFORMA = 7
    stISSUED = 3
    stCANCELLED = 6
    stBACKUP = 9
End Enum

Public Enum enSortField
    enSequence = 1
    enTitle = 2
    enAuthor = 3
    enCode = 4
    enRef = 5
End Enum

Public Enum enumReportPresentation
    enPreview = 1
    enPrintout = 2
    enCSV = 3
End Enum


Public Enum enumCommandType
    enTable = adCmdTable
    enQuery = adCmdStoredProc
    enText = adCmdText
    enStoredProcedure = adCmdStoredProc
End Enum

Public Enum enumWildcard
    enNone = 0
    enEnd = 1
    enStart = 2
    enBoth = 3
End Enum
Public Enum enumActionTaken
    enFulfilled = 9
    enPOIssued = 1
    enNoAction = 0
End Enum

Public Enum enumCOType
    enNormalCO = 1
    enWant = 2
    enStanding = 3
End Enum
'Public Enum enumTRType
'    enOrder = 6
'    enAppro = 9
'    enInvoice = 3
'End Enum

Public Enum enCatalogue
    enGetsCatalogueYes = 1
    enGetsCatalogueNo = 2
    enGetsCatalogueEither = 3
End Enum
Public Enum enOverseas
    enOverseasYes = 1
    enOverseasNo = 2
    enOverseasEither = 3
End Enum
Public Enum enMailType
    enAirmail = 1
    enSurfaceMail = 2
    enEitherMail = 3
    enAll = 4
End Enum

Public Enum enumSince
    enDay = 0
    enAny = 1
    enWeek = 2
    enMonth = 3
    enQuarter = 4
    enYear = 5
End Enum
Public Enum enProductSTatus
    enInPrint = 0
    enOutOfPrint = 1
    enAwaitingReprint = 2
    enOnBackorder = 3
    enMarketRestricted = 4
    enNotYetPrinted = 5
End Enum

Public Enum enUseStatus
    enActive = 0
    enInActive = 1
End Enum
Public Enum enProductDataSource
    enLocalDB = 0
    enBF = 1
End Enum

Public Enum enTransmitType
    enPrint = 1
    enEDI = 2
    enMail = 3
    enView = 4
    enXML = 5
End Enum

Public Enum enumExportType
    enProduct = 1
    enDocument = 2
    enSupplier = 3
    encustomer = 4
End Enum



 Enum enumSecurityNode
    enSECURITY_PO_SIGN = 1
    enSECURITY_RETREQ_SIGN = 2
    enSECURITY_APPR_SIGN = 3
    enSECURITY_TFR_SIGN = 4
    enSECURITY_CN_SIGN = 5
    enSECURITY_RETFIN_SIGN = 6
    enSECURITY_INV_SIGN = 7
    enSECURITY_CO_SIGN = 8
    enSECURITY_APP_SIGN = 9
    enSECURITY_STKADJ_SIGN = 10
    enSECURITY_CASHUP_SIGN = 11
    enSECURITY_CONFIG_SIGN = 12
    enSECURITY_DICT_SIGN = 13
    enSECURITY_GRN_SIGN = 14
    enSECURITY_COMM_AUTH = 15
    enSECURITY_CUSTDISCOUNT_AUTH = 16
    enSECURITY_EDITDICT_AUTH = 17
    enSECURITY_EDITPRODUCTTYPES_AUTH = 18
    enSECURITY_EDITROUNDINGRULES_AUTH = 19
    enSECURITY_MERGESTOCK_AUTH = 20
    enSECURITY_MERGECUSTSUPP_AUTH = 21
    enSECURITY_ISSUPERVISOR = 22
    enSECURITY_ISREP = 23
    enSECURITY_ISOPERATOR = 24
    enSECURITY_ACTIVE = 25
    enSECURITY_MULTIBUY = 26
    enSECURITY_ACCEPTDIRECTDEPOSIT = 27
    enSECURITY_VOIDEXCHANGE = 28
    enSECURITY_POSDISCOUNT = 29
    enSECURITY_ISSUEPOSAPPRO = 30
    enSECURITY_POSPRICECHANGE = 31
    enSECURITY_ISSUEPOSREFUND = 32
    enSECURITY_ISSUEPOSCREDITNOTE = 33
    enSECURITY_REFUNDDEPOSITONPOS = 34
    enSECURITY_ACCEPTACPAYMENT = 35
    enSECURITY_QU_SIGN = 36
    enSECURITY_BLOCK_DEBTORS = 37
    enSECURITY_OPENDRAWER = 38
    enSECURITY_ISSUEPETTYCASH = 39
    enSECURITY_CREATENEWSTOCKITEM = 40
    enSECURITY_SALELINEDELETE = 41
    
End Enum



Public Enum enumAuditPoint
    enProductPriceDiscount = 1
    enCustomerDiscountTerms = 2
End Enum

Public Enum enumDocType
    enTypeInvoice = 3
    enTypePurchaseOrder = 5
    enTypeCustomerOrder = 1
    enTypeAppro = 2
    enTypeApproReturn = 10
    enTypeReturn = 11
    enTypeCreditNote = 8
    enTypeDelivery = 10
End Enum
Public Enum OlDefaultFolders
  olFolderDeletedItems = 3
  olFolderOutbox = 4
  olFolderSentMail = 5
  olFolderInbox = 6
  olFolderCalendar = 9
  olFolderContacts = 10
  olFolderJournal = 11
  olFolderNotes = 12
  olFolderTasks = 13
  olFolderDrafts = 16
  olPublicFoldersAllPublicFolders = 18
End Enum
Public Enum enumCustomerType
    enPrivate = 1
    enBookclub = 2
    enBusiness = 3
    enLoyalty = 4
End Enum


Public Function TranslateAuditPoint(val As Integer) As String
    Select Case val
    Case 1
        TranslateAuditPoint = "Product edit"
    Case 2
        TranslateAuditPoint = "Customer terms and disc."
    Case Else
        TranslateAuditPoint = "Unknown"
        
    End Select
End Function



