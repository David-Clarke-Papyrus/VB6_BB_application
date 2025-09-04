Attribute VB_Name = "UDTPOS"
Public Type ZSessionProps
    ID As String * 40
    StartDate As Date
    EndDate As Date
    NominalDate As Date
    SMID As Long
    TotalValueSales As Long
    TotalValueCredit As Long
    TotalValueDiscount As Long
    Cash As Long
    GiftVouchers As Long
    Cheques As Long
    CCVouchers As Long
    TillPoint As String * 40
    SupervisorName As String * 20
    Reportable As Integer
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type ZSessionData
     buffer As String * 150
End Type

'OPSSessions are actually CS's so look there for the code

Public Type ExchangeProps
    ID As String * 40
    ZID As String * 40
    OPSID As String * 40
    ExchangeDate As Date
    SupervisorID As Long
    OperatorID As Long
    TotalPayable As Long
    TotalCredit As Long
    TotalDiscount As Long
    TotalExtras As Long
    ChangeGiven As Long
    TotalPayment As Long
    BalanceOwing As Long
    LoyaltyValue As Long
    TotalVAT As Long
    TotalExVAT As Long
    TotalExtension As Long
    TotalQty As Long
    TPID As Long
    ToVoid As Long
    Status As String * 5
    OR_ActionedDate As Date
    Type As String * 4
    Note As String * 5000
    ExchangeCode As String * 15
    SalesPersonName As String * 20
    SalesRepID As Long
    DiscountRate As Double
    ExchangeNumber As Long
    ShowVAT As Boolean
    VATable As Boolean
    StaffName As String * 30
    DocumentCode As String * 20
    IdentifyCustomer As Boolean
    VOIDED As Boolean
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type ExchangeData
     buffer As String * 5274
End Type

Public Type PaymentProps
    PaymentID As Long
    Amt As Long
    AmtTendered As Long
    COLID As Long
    Type As String * 4
    Reference As String * 100
    Note As String * 50
    EXCHANGE_GUID As String * 40
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type

Public Type PaymentData
     buffer As String * 206
End Type

Public Type SaleProps
    ID As Long
    EXCHANGE_GUID As String * 40
    Code As String * 20
    CodeF As String * 20
    PID As String * 40
    title As String * 40
    Author As String * 40
    Counterfoil As String * 30
    DiscountRule As String * 30
    COLID As Long
    Qty As Long
    Price As Long
    TempPrice As Long
    PriceAlteration As Long
    VAT As Double
    Discount As Long
    Payable As Long
    DiscountRate As Double
    VATRate As Double
    Note As String * 80
    Ref As String * 25
    DiscountDescription As String * 20
    Sequence As Long
    LoyaltyRate As Long
    ServiceItem As Boolean
    CatID As Long
    ActionSignatureID As Long
    MBCode As String * 15
    IdentifyCustomer As Boolean
    NoDiscountAllowed As Boolean
    IsDepositItem As Boolean
    SaleGUID As String * 40
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type

Public Type SaleData
     buffer As String * 491
End Type

Public Type CustomerPOSProps
    ID  As Long
    CustomerTypeID As String * 100
    DefaultDiscount As Double
    Balance As Long
    CreditLimit As Long
    Name As String * 100
    Initials As String * 15
    title As String * 10
    AcNo As String * 15
    SoundexName As String * 10
    Phone As String * 25
    SearchPhone As String * 10
    Note As String * 500
    VATable As Boolean
    Address As String * 100
    Balances As String * 50
    FullIdentification As String * 120
    CType As String * 10
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type CustomerPOSData
     buffer As String * 1084
End Type

Public Type COLPOSProps
    COLID  As Long
    TPID As Long
    TRID As Long
    COLDate As Date
    Code As String * 15
 '   DocCode As String * 15
    PID As String * 40
    Description As String * 40
    Qty As Long
    QtyDispatched As Long
    Price As Long
    DiscountRate As Double
    Deposit As Long
    DepositStatus As String * 1
End Type
Public Type COLPOSData
     buffer As String * 119
End Type
Public Type APPPOSProps
    APPID As Long
    TPID As Long
    DOCDate As Date
    DocCode As String * 15
End Type
Public Type APPPOSData
     buffer As String * 24
End Type

Public Type APPLPOSProps
    APPID As Long
    APPLID As Long
    TPID As Long
    title As String * 100
    Author As String * 30
    CodeF As String * 20
    Code As String * 15
    QtyOut As Long
    QtyBack As Long
    Price As Long
    QtyNett As Long
    DiscountRate As Double
    VATRate As Double
    Total As Long
    PID As String * 40
    
    DocCode As String * 15
End Type
Public Type APPLPOSData
     buffer As String * 246
End Type

Public Type InvoicesContainingProps
    TRID As Long
    DocCode As String * 30
    DOCDate As Date
    DocValue As Long
    PID As String * 40
End Type

Public Type InvoicesContainingData
    buffer As String * 78
End Type

Public Type InvoiceLinesProps
    Description As String * 50
    LineCode As String * 20
    Qty As Long
    QtyCredited As Long
    Price As Long
    DiscountRate As Double
    SalesValue As Long
    ILID As Long
    PID As String * 40
End Type

Public Type InvoiceLinesData
    buffer As String * 124
End Type
Sub lenExchangeProps()
Dim x As CustomerPOSProps
    MsgBox LenB(x) & "        " & LenB(x) / 2
End Sub


