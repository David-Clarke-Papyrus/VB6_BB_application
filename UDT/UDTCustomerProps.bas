Attribute VB_Name = "UDTCustomerProps"
Option Explicit
Public Type tRule
    Criterion As String * 200
    Operator As String * 50
    Argument As String * 1200
    Description As String * 100
    ID As Long
End Type
Public Type CustomerProps
    CustID  As Long
    DefaultAddressID As Long
    Role As Integer
    ParentCustomerID    As Long
    ParentCustomerName As String * 30
    CustomerTypeID As Long
    DefaultDiscount As Double
    CreditLimit As Long
    DefaultDeliveryDays As Long
    Name As String * 100
    Initials As String * 15
    Title As String * 10
    MOBILE As String * 50
    AcNo As String * 15
    SAN As String * 20
    AccAcno As String * 15
    IDNUM As String * 20
    SoundexName As String * 10
    Phone As String * 50
    SearchPhone As String * 15
    StoreName As String * 30
    StoreID As Long
    CustomerCategory As Integer
    Note As String * 500
    VatNumber As String * 25
    ShowVAT As Boolean
    PaymentStyle As String * 1
    OurACnoWithClient As String * 50
    Terms As Integer
    BalCur As Double
    Bal30 As Double
    Bal60 As Double
    Bal90 As Double
    Bal120 As Double
    Bal120Plus As Double
    Balance As Double
    Blocked As Boolean
    oLineToInvoice As Boolean
    UseQuotedPrice As Boolean
    VATable As Boolean
    GetsCatalogue As Boolean
    CanBeDeleted As Boolean
    OneLinePerInvoice As Boolean
    CompleteOrder As Boolean
    CustNotifyBookLaunch As Boolean
    CustNotifyBookSale As Boolean
    CustNotifyBookPromotion As Boolean
    DispatchMethod As String * 1
    DateRecordAdded As Date
    DateLastModified As Date
    SalesRepID As Long
    Repname As String * 20
    SalesOrderTemplateName As String * 12
    QuotationTemplateName As String * 12
    ApproTemplateName As String * 12
    ApproReturnTemplateName As String * 12
    InvoiceTemplateName As String * 12
    GDNTemplateName As String * 12
    CreditNoteTemplateName As String * 12
    ContactPerson As String * 50
    ContactpersonPhone As String * 50
    OpenItemOrBalBF As String * 1
    GenerateSeparateInvoicesForSeparateOrders As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type CustomerData
     buffer As String * 1270
End Type

Public Type IGProps
    IGID As Long
    TPID As Long
    Description As String * 30
    IsLoyalty As Boolean
    IsLaunch As Boolean
    ISPromotion As Boolean
    IsSale As Boolean
    IsLitLunch As Boolean
    IsBookClub As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type IGData
     buffer As String * 45
End Type

Public Type CustomerPropsDisplay
    ID  As Long
    Initials As String * 15
    Appell As String * 15
    Name As String * 50
    AcNo As String * 12
    CELL As String * 18
    L1 As String * 50
    L2 As String * 50
    L3 As String * 50
    L4 As String * 50
    L5 As String * 50
    L6 As String * 50
    Country As String * 25
    PostCode As String * 18
    Addressee As String * 70
    FullIdentification As String * 80
    Phone As String * 30
    CustomerTypeID As Long
    CustomerTypeDescription As String * 15
    MailingAddress As String * 200
    EMail As String * 100
    DefaultAddressID As Long
    GetsCatalogue As Boolean
    SalesValue As Double
    SalesQty As Double
    DOB As Date
    Note As String * 500
    Temporary As Boolean
    Blocked As Boolean
    Balance As Double
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type CustomerDataDisplay
     buffer As String * 1480
End Type
Public Type SPCProps
    Price As Long
    dateOfSale As Date
    Week As Integer
    Qty As Long
    Valu As Long
    Title As String * 120
    code As String * 15
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type SPCData
     buffer As String * 150
End Type
Public Type DocsTPProps
    TRID As Long
    TRDATE As Date
    TRCODE As String * 120
    Type As Integer
    TRSTATUS As Integer
    TRValue As Long
    OrderType As Integer
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type DocsTPData
     buffer As String * 150
End Type
Public Type DebtorsProps
    TPID As Long
    TRID As Long
    REMID As Long
    TRDATE As Date
    TRCaptureDATE As Date
    TRProcessingDate As Date
    TRCODE As String * 20
    CreditorDocDate As Date
    DueDate As Date
    SettlementDueDate As Date
    Credit As Double
    PayableAmount As Double
    PayableAmountF As String * 30
    PayableAfterSettDisc As Double
    PayableAfterSettDiscF As String * 30
    EffectivePayableAfterSettDiscF As String * 30
    EffectivePayableAfterSettDisc As Double
    ClaimValue As Double
    ClaimValueF As String * 30
    SupplierInvoiceCode As String * 30
    TempRemittance As Double
    TempRemittanceF As String * 30
    TempPayment As Double
    Debit As Double
    Owing As Double
    OwingF As String * 30
    VATAmount As Double
    DocType As String * 25
    Memo As String * 300
    Status As String * 10
    BFTotal As Double
    BFCur As Double
    BF30 As Double
    BF60 As Double
    BF90 As Double
    BF120 As Double
    
    dbDoc As String * 15
    dbDate As String * 20
    dbDocType As String * 3
    dbAmt As String * 20
    crDoc As String * 15
    crDate As String * 20
    crDocType As String * 3
    crAmt As String * 20
    crTotal As String * 20
    Balance_OI As String * 20
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type DebtorsData
     buffer As String * 820
End Type

Public Type RemittanceProps
    TPID As Long
    TRID As Long
    CustomerAcno As String * 20
    CustomerName As String * 100
    DocumentCode As String * 20
    DocumentNominalDate As Date
    DocumentNumberDate As Date
    DocumentIssueDate As Date
    DocumentOpenStatus As Boolean
    DocumentStatus As Long
    DocumentNote As String * 100
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type RemittanceData
     buffer As String * 264
End Type


Public Type TPAttributesProps
    AttributeID As Long
    Name As String * 30
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type TPAttributesData
     buffer As String * 40
End Type

Sub testTPATTRIB()
Dim x As IGProps
    MsgBox LenB(x) & "    " & LenB(x) / 2
End Sub
