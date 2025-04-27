Attribute VB_Name = "UDTInvoice"
Public Type InvoiceProps
    InvoiceID As Long
    ID As String * 40
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
    TotalDiscountExVAT As Long
    TotalnonExtrasExVAT As Long
    TotalDiscount_Foreign As Long
    TotalExtension As Long
    TotalExtension_Foreign As Long
    TotalNonVAT As Long
    TotalNonVAT_Foreign As Long
    TotalVAT As Long
    VATRoundingAdjustment As Double
    TotalVATLineSummed As Long
    TotalVAT_Foreign As Long
    TotalExtras As Long
    TotalExtras_Foreign As Long
    TotalPayable As Long
    TotalPayable_Foreign As Long
    TotalServiceItem As Long
    TotalServiceItem_Foreign As Long
    Exchange As String * 40
    TotalQty As Long
    StaffID As Long
    StaffName As String * 10
    Signature As String * 50
    StaffEmail As String * 50
    TPName As String * 100
    TPPhone As String * 25
    TPFax As String * 25
    TPACCNum As String * 14
    Memo As String * 200
    ForAttn As String * 100
    Waybill As String * 100
    CourierURL As String * 90
    BusPhone As String * 25
    DOCCode As String * 14
    CurrencyFormat As String * 14
    Log As String * 500
    DOCDate As Date
    CaptureDate As Date
    ProcessingDate As Date
    CurrencyFactor As Double
    VATRate As Double
    DiscountRate As Double
    DepositPaid As Boolean
    ShowVAT As Boolean
    NonVATDocument As Boolean
    VATable As Boolean
    STATUS As Integer
    Proforma As Boolean
    IsPreDelivery As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type InvoiceData
    Buffer As String * 1514
End Type

Public Type ILProps
    Sequence As Long
    InvoiceLineID As Long
    InvoiceID As Long
    PIID As Long
    Price As Long
    PriceExVAT As Long
    Cost As Long
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
    DistributorName As String * 30
    DistributorAcno As String * 10
    Note As String * 350
    code As String * 15
    EAN As String * 13
    PID As String * 40
    GDNCode As String * 20
    COLID As Long
    APPLID As Long
    APPLQTY As Long
    CreditedQty As Long
    MainAuthor As String * 40
    Title As String * 120
    Article As String * 10
    TitleWithArticle As String * 130
    CodeForExport As String * 18
    CodeF As String * 20
    VATRate As Double
    DiscountPercent As Double
    Serial As Integer
    tmpCNLQty As Integer
    Marker As Boolean
    ServiceItem As Boolean
    SubstitutesAvailable As Boolean
    SubstitutionForPID As String * 40
    CO_StaffShortname As String * 5

    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type ILData
    Buffer As String * 980
End Type
Public Type dInvoiceProps
    TRID As Long
    Qty As Long
    QtyCredited As Long
    PIID As Long
    ILID As Long
    Price As Long
    STATUS As Integer
    CURRID As Long
    StaffID As Long
    CurrRate As Double
    TPName As String * 100
    DOCCode As String * 14
    TPAccNo As String * 10
    Discount As Double
    VATRate As Double
    Title As String * 30
    Log As String * 500
    SM As String * 10
    DOCDate As Date
    CaptureDate As Date
    Proforma As Boolean
    Exchange As String * 40
    IsPreDelivery As Boolean
    InvoiceValue As Double
    InvoiceQty As Long
    CustomerDisplay As String * 100
End Type
Public Type dInvoiceData
    Buffer As String * 850
End Type
Sub G1()
      Dim x As ILProps
42310 MsgBox LenB(x) / 2
End Sub

