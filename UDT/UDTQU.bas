Attribute VB_Name = "UDTQU"
Public Type QUProps
    QuoteID As Long
    BillToAddressID As Long
    DelToAddressID As Long
    TPID As Long
    SalesPersonID As Long
    SalesRepID As Long
    SalesRepName As String * 20
    CustPaid As Boolean
    CommPaid As Boolean
    COMPID As Long
    CurrencyID_Foreign As Long
    Postage As Long
    Insurance As Long
    TotalDiscount As Long
    TotalDiscount_Foreign As Long
    TotalExtension As Long
    TotalExtension_Foreign As Long
    TotalNonVAT As Long
    TotalNonVAT_Foreign As Long
    TotalVAT As Long
    TotalVAT_Foreign As Long
    TotalExtras As Long
    TotalExtras_Foreign As Long
    TotalPayable As Long
    TotalPayable_Foreign As Long
    TotalServiceItem As Long
    TotalServiceItem_Foreign As Long
    TotalQty As Long
    StaffID As Long
    StaffName As String * 10
    Signature As String * 50
    StaffEmail As String * 50
    TPNAME As String * 100
    TPPhone As String * 25
    TPFax As String * 25
    TPACCNum As String * 14
    Memo As String * 200
    ForAttn As String * 100
    BusPhone As String * 25
    DOCCode As String * 14
    CurrencyFormat As String * 14
    Log As String * 500
    Ref As String * 50
    OrderDate As Date
    DOCDate As Date
    CaptureDate As Date
    CurrencyFactor As Double
    VATRate As Double
    DiscountRate As Double
    DepositPaid As Boolean
    ShowVAT As Boolean
    VATable As Boolean
    Status As Integer
    Proforma As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type QUData
    buffer As String * 1288
End Type

Public Type QULProps
    Sequence As Long
    QULID As Long
    QUID As Long
    Price As Long
    PriceExVAT As Long
    SalesValue As Long
    VATAmount As Long
    ForeignPrice As Long
    FCFactor As Double
    FCID As Long
    Qty As Long
    QtyFirm As Long
    QtySS As Long
    Deposit As Long
    ForeignDeposit As Long
    Ref As String * 30
    Publisher As String * 20
    Note As String * 350
    code As String * 15
    EAN As String * 13
    PID As String * 40
    COLID As Long
    MainAuthor As String * 40
    Title As String * 120
    CodeForExport As String * 18
    CodeF As String * 20
    VATRate As Double
    ExtraVATRate As Double
    DiscountPercent As Double
    Serial As Integer
    Marker As Boolean
    ServiceItem As Boolean
    SubstitutesAvailable As Boolean
    SubstitutionForPID As String * 40
    ExtraPID As String * 40
    ExtraCode As String * 20
    ExtraCharge As Long
    ExtraChargeDescription As String * 50
    ID As String * 40
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type QULData
    buffer As String * 914
End Type
Public Type dQUProps
    TRID As Long
    Qty As Long
    QULID As Long
    Price As Long
    Status As Integer
    CURRID As Long
    StaffID As Long
    CurrRate As Double
    TPNAME As String * 100
    DOCCode As String * 14
    TPAccNo As String * 10
    CustomerDisplay As String * 100
    Discount As Double
    VATRate As Double
    Title As String * 30
    Log As String * 500
    SM As String * 10
    DOCDate As Date
    CaptureDate As Date
    Proforma As Boolean
End Type
Public Type dQUData
    buffer As String * 802
End Type

Sub qu()
Dim x As QULProps
MsgBox LenB(x) / 2
End Sub

