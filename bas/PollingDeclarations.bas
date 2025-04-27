Attribute VB_Name = "Declarations"
Option Explicit
'Global oADODBConn As ADODB.Connection
'Global oConfig As a_Configuration
Global oPC As z_POSConnection
'Global oError As a_Error
'Global oHost As a_Host
'Global oBookfind As a_BookFind
Public Const MEMOLENGTH As Integer = 4096
Public tmpErr As String
Public tmpError As String
Public lngResult As Long
Public retval
'Public oCurrencyManager As z_CurrencyManager
Dim i As ListBox

Global Const PUBLISHER_ROLE_ID = 1
Global Const SUPPLIER_ROLE_ID = 2
Global Const CUSTOMER_ROLE_ID = 3
Global Const OWNER_ROLE_ID = 5
Global Const TRANSTYPE_STOCKTAKE = 7
Global Const TRANSTYPE_CUSTOMER_ORDER = 6
Global Const TRANSTYPE_SUPPLIER_ORDER = 5
Global Const TRANSTYPE_DELIVERY = 4
Global Const TRANSTYPE_INVOICE = 3
Global Const TRANSTYPE_CASHSALE = 2
Global Const TRANSTYPE_TRANSFER = 1
Global Const TRANSTYPE_RETURN = 8
Global Const TRANSTYPE_APPRO = 9
Global Const TRANSTYPE_APPRORETURN = 10
Global Const NEW_BOOKS = "NB"
Global Const CUSTOMER_ORDER = "CO"
Global Const STOCK_ORDER = "ST"


Public Type ConfigProps
    ClientName As String * 50
    ApplicationName As String * 50
    TransactionPrefix As String * 3
    Version As String * 10
    WSStart As Date
    MSStart As Date
    BaseYear As Integer
    StockTakeDir As String * 150
    CashSaleDir As String * 150
    TransferDir As String * 150
    PsionDir As String * 150
    DatabaseName As String * 255
    ErrorLogDir As String * 150
    HelpFileName As String * 150
    VATRate As Double
    VatNumber As String * 20
    fConfirmBeforePrint As Boolean
    fPrintPrices As Boolean
    fPrintDiscount As Boolean
    fPrintMemo As Boolean
    ClientsLogo As Object
    ApplicationLogo As Object
    LocalCurrencyID As Long
    OwnerID As Long
    DefaultCompany As Long
    LastUpdateDate As Date
    nextUpdateTime As Date
    BookfindPassword As String * 20
    OrderText As String * 255
    LastStockTakeDate As Date
    fAllowCopyInfo As Boolean
    CompanyRegistration As String * 20
    fTrimTitles As Boolean
    OfferSignature As String * 20
    CustomerOrderReport As String * 25
    BookReport As String * 25
    fShowDeposit As Boolean
    InvoiceReport As String * 25
    InvoiceCopies As Integer
    DeliveryReport As String * 25
    CurrentCatalogID As Long
    RandDollar As Double
    RandSterling As Double
    Banner As String * 255
    Approreport As String * 25
    ApproCopies As Integer
    ApproReturnReport As String * 25
    ApproreturnCopies As Integer
    LastCatalogDoc As String * 50
    CatalogDir As String * 150
    CatalogTemplateDir As String * 150
    OrderReportname As String * 25
    OrderCopies As Integer
    AntInvoice As String * 25
    AntInvoiceCopies As Integer
    DaysInWeek As Integer
    OpSetsAuto As Long
    PatchNumber As Long
    LastWantsExportDate As Date
End Type
Public Type ConfigData
    Buffer As String * 4734
End Type

Public Type ProductProps
    ID As Long
    Code As String * 10
    EAN As String * 13
    Availability As String * 5
    Description As String * 300
    CategoryID As Long
    BFClassification As String * 10
    UKPrice As Currency
    UKPoundsExch As Double
    USPrice As Currency
    USDollarExch As Double
    LastExchUpdate As Date
    LocalPrice As Currency
    Cost As Currency
    Title As String * 255
    BindingCode As String * 5
    SeriesTitle As String * 255
    SubTitle As String * 255
    Author As String * 255
    Publisher As String * 50
    Note As String * 40
    PublisherID As Long
    StockBalance As Long
    PublicationDate As String * 50
    PublicationPlace As String * 100
    MainSupplierName As String * 30
    Edition As String * 100
    LastSupplierID As Long
    LastDEalID As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type ProductData
     Buffer As String * 3478
End Type
Public Type CSProps
    ID As Long
    TRID As Long
    TRTPID As Long
    Code As String * 10
    DAteStarted As Date
    DateIssued As Date
    TextFileFullPathAndName As String * 255
    FrontDeskComputerName As String * 255
    TAType As Integer
    OperatorID As Long
    OperatorName As String * 30
    Note As String * 255
    Issued As Boolean
    Void As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type CSData
    Buffer As String * 1656
End Type

Public Type CSLProps
    ID As Long
    CSID As Long
    TRID As Long
    QTY As Integer
    DateTime As Date
    ProductID As Long
    EAN As String * 13
    EANFormatted As String * 16
    ISBN As String * 10
    Title As String * 40
    Author As String * 20
    Price As Currency
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type CSLData
    Buffer As String * 244
End Type

'For search Tool-----------
Public Type SearchProps
  BookID As Long
  ISBN As String * 15
  Title As String * 50
  Author As String * 50
  Price As Double
  Stock As Long
  Publisher As String * 50
End Type
Public Type BookSearchData
    Buffer As String * 348
End Type
'--------------------------

Public Type CurrencyProps
  ID  As Long
  Symbol As String * 1
  Format As String * 15
  Divisor As Double
  ConversionToLocalFactor As Double
  Description As String * 15
End Type
  
'---------------------------

Public Type OperatorDisplayProps
    ID As Long
    Fullname As String * 40
End Type
Public Type OperatorDisplayData
    Buffer As String * 84
End Type
Public Type TextListProps
    Key As String * 30
    Item As String * 255
End Type
Public Type TextListData
    Buffer As String * 285
End Type
Public Type ProductDisplayProps
    ID As Long
    Description As String * 255
    Author As String * 50
End Type
Public Type ProductDisplayData
    Buffer As String * 60
End Type

Public Sub TestLength()
    Dim t As ProductProps
    MsgBox LenB(t)
End Sub


