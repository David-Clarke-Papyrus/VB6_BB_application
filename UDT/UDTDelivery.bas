Attribute VB_Name = "UDTDelivery"
Option Explicit

Public Type DELProps
    TPID As Long
    TRID As Long
    CurrencyID As Long
    CurrencyRate As Double
    COMPID As Long
    TotalExtension As Long
    TotalExtensionShort As Long
    TotalVAT As Long
    TotalDiscount As Long
    TotalPayable As Long
    TotalExtensionSimple As Long
    TotalExtensionSimple_Foreign As Long
    TotalExtension_Foreign As Long
    TotalExtensionShort_Foreign As Long
    TotalVAT_Foreign As Long
    TotalDiscount_Foreign As Long
    TotalPayable_Foreign As Long
    Status As Integer
    DOCDate As Date
    CaptureDate As Date
    ProcessingDate As Date
    DOCCode As String * 30
    Memo As String * 100
    TPNAME As String * 100
    TPACCNum As String * 14
    SupplierInvoiceRef As String * 50
    SupplierInvoiceDate As Date
    BatchTotal As Long
    BatchQtyTotal As Long
    BatchTotalExtras As Long
    TotalQtyItems As Long
    TotalQtyShort As Long
    StaffID As Long
    StaffName As String * 10
    VATable As Boolean
    VATRate As Double
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type DELData
    buffer As String * 380
End Type

Public Type DELLProps
    DELLID As Long
    TRID As Long
    PIID As Long
    POLID As Long
    COLID As Long
    QtyFirm As Long
    QtySS As Long
    QtyShort As Long
    QtyTotal As Long
    Price As Long
    CorrectedPrice As Long
    ForeignPrice As Long
    CorrectedForeignPrice As Long
    PriceSell As Long
    Cost As Long
    DiscountedPrice As Long
    VATRate As Double
    Discount As Double
    CorrectedDiscount As Double
    Ref As String * 20
    PID As String * 40
    CodeFForExport As String * 20
    code As String * 15
    CodeF As String * 20
    EAN As String * 13
    Title As String * 255
    Section As String * 51
    Note As String * 50
    Status As Integer
    PO_Code As String * 12
    POL_Price As Long
    POL_QtySS As Long
    POL_QtyFirm As Long
    POL_ForeignPrice As Long
    ProductTypeID As Long
    PO_CURRID As Long
    ClaimID    As Long
    PO_CURRRATE As Long
    POL_Discount As Double
    ReasonID As String * 21
    MBCode As String * 10
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type DELLData
    buffer As String * 596
End Type

Public Type dDELProps
    TRID As Long
    TPID As Long
    DELID As Long
    QtyTotal As Long
    PIID As Long
    DELLID As Long
    Price As Long
    StaffID As Long
    Discount As Double
    PID As String * 40
    TPNAME As String * 100
    TPAccNo As String * 10
    Title As String * 30
    DOCDate As Date
    DOCCode As String * 14
    CaptureDate As Date
    InvoiceRef As String * 50
    InvoiceDate As Date
    InvoiceValueF As String * 20
    InvoiceValue As Double
    InvoiceQty As Long
    InvoiceShort As Long
    Status As String * 15
    DeliveredValue As Double
    DeliveredQty As Double
    CURRID As Long
    
End Type
Public Type dDELData
    buffer As String * 310
End Type

Public Type dDELLProps
    DELLID As Long
    TRID As Long
    Qty As Long
    Price As Long
    ForeignPrice As Long
    PriceSell As Long
    Discount As Double
    SINVRef As String * 50
    SINVDate As String * 15
    PID As String * 40
    code As String * 15
    CodeF As String * 20
    DocRef As String * 10
    DOCDate As Date
    Title As String * 255
    Status As Integer
    CurrDivisor As Integer
    CurrFormatString As String * 15

    CURRID As Long
    CurrRate As Long

    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type dDELLData
    buffer As String * 450
End Type



Sub DEL()
Dim x As dDELLProps
    MsgBox LenB(x) & "     " & LenB(x) / 2
End Sub
