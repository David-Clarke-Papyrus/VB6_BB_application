Attribute VB_Name = "UDTCN"
Public Type CNProps
    COMPID As Long
    TRID As Long
    CustomerID As Long
    BillToAddressID As Long
    DelToAddressID As Long
    StaffID As Long
    TotalDiscount  As Long
    TotalNonVAT As Long
    TotalNonVAT_Foreign As Long
    TotalVAT As Long
    TotalPayable As Long
    TotalExtension As Long
    TotalDiscount_Foreign  As Long
    TotalVAT_Foreign As Long
    TotalPayable_Foreign As Long
    TotalExtension_Foreign As Long
    TotalQty As Long
    ForeignCurrencyID As Long
    CurrRate As Double
    OrderNum As String * 10
    DOCDate As Date
    CaptureDate As Date
    TPID As Long
    TPNAME As String * 100
    TPPhone As String * 25
    TPFax As String * 25
    TPACCNum As String * 14
    TPMemo As String * 200
    BusPhone As String * 25
    VerificationNum As String * 20
    CardNum As String * 40
    ExpiryDate As String * 20
    PaymentMethod As String * 30
    DOCCode As String * 14
    StaffName As String * 10
    Log As String * 500
    
    Amount As Currency
    VATable As Boolean
    ShowVAT As Boolean
    Status As Integer
    Memo As String * 200
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type CNData
    buffer As String * 1300
End Type

Public Type CNLProps
    Sequence As Long
    CNLineID As Long
    TRID As Long
    CustID As Long
    PID As String * 40
    PIID As Long
    Qty As Long
    QtyDAM   As Long
    InvPrice As Long
    ServiceItem As Boolean
    VATValue As Long
    InvForeignPrice As Long
    DiscountRate As Double
    INVLineID As Long
    INVLineCode As String * 10
    InvLineDate As Date
    InvLineRef As String * 25
    Title As String * 120
    Author As String * 30
    Publisher As String * 20
    ProductCodeForExport As String * 20
    ProductCodeF As String * 20
    ProductCode As String * 20
    EAN As String * 13
    Note As String * 50
    VATRate As Double
    Marker As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type

Public Type CNLData
    buffer As String * 455
End Type


Public Type dCNProps
    COMPID As Long
    TRID As Long
    StaffID As Long
    TPNAME As String * 50
    TPAccNo As String * 10
    Phone As String * 25
    DOCCode As String * 13
    CustomerDisplay As String * 100
    DOCDate As Date
    CaptureDate As Date
    CustRef As String * 20
    Status As String * 20
End Type
Public Type dCNData
    buffer As String * 252
End Type

Sub GetCNLen()
Dim x As dCNProps
MsgBox LenB(x)
MsgBox LenB(x) / 2
End Sub

