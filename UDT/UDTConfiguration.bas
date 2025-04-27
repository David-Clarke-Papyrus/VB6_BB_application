Attribute VB_Name = "UDTConfiguration"
Public Type ConfigProps
    TransactionPrefix As String * 3
    VatNumber As String * 20
    UsesBookfind As Boolean
    ShowProdsWithInstancesOnly As Boolean
    UsesGardners As Boolean
    OrderText As String * 500
    QuotationText As String * 500
    SalesOrderText As String * 500
    InvoiceText As String * 500
    StatementText As String * 500
    EmailPOMsg As String * 500
    EmailInvMsg As String * 500
    EmailCNMsg As String * 500
    EmailAPPMsg As String * 500
    EmailQuoteMsg As String * 500
    
    OfferSignature As String * 20
    LookupSeq As String * 10
    PrintingSettings As String * 500
    COLAllocationStyle As String * 1
    GFXNumber As String * 20
    IE_AccountingApplicationName As String * 50
    IE_CreditorsContraAccount As String * 50
    IE_DebtorsContraAccount As String * 50
    IE_GLReference As String * 50
    TFRDiscount As Integer
    COTypesSupported As Integer
    MinMU As Integer
    LoyStartNumber As Long
    LoyEndNumber As Long
    VATRate As Double
    ConfirmBeforePrint As Boolean
    PrintPrices As Boolean
    DiscountVATYN As Boolean
    CasualCustomersYN As Boolean
    PrintMemo As Boolean
    AllowCopyInfo As Boolean
    AntiquarianYN As Boolean
    NonBookYN As Boolean
    SignTransactions As Boolean
    EnforceSections As Boolean
    ShowDeposit As Boolean
    IsVATRegion As Boolean
    DiscountVAT As Boolean
    ReorderPerCOL As Boolean
    AggregatePOs As Boolean
    EnforceCOLRef As Boolean
    SupportsWORD As Boolean
    CaptureDecimal As Boolean
    SupportsLoyaltyClub As Boolean
    LocalCountryID As Long
    DefaultCurrencyID As Long
    LocalCurrencyID As Long
    DefaultCOMPID As Long
    DefaultStoreID As Long
    BillToStoreID As Long
    DefaultPT As Long
    DefaultCT As Long
    DefaultSection As Long
    CSCustomerID As Integer
    CustomerTypeDictID As Integer
    SectionTypeDictID As Integer
    BusinessCustomerTypeID As Integer
    PrivateCustomerTypeID As Integer
    BookClubCustomerTypeID As Integer
    InterestGroupsDictID As Integer
    LoyaltyClubType As Integer
    TFRDiscountAdj As Integer
    LastStockTakeDate As Date
    LastWantsExportDate As Date
    LastUpdateDate As Date
    StatementAsAt As Date
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type ConfigData
    buffer As String * 5860
End Type

Sub configLen()
Dim x As ConfigProps
    MsgBox LenB(x) & "   " & LenB(x) / 2
End Sub
