Attribute VB_Name = "UDTReturn"
Option Explicit

Public Type RProps
    TRID As Long
    TPID As Long
    DOCCode As String * 10
    ApprovalRef As String * 20
    ApprovalTermDate As Date
    DOCDate As Date
    CaptureDate As Date
    TPNAME As String * 100
    TPACCNum As String * 14
    StaffID As Long
    CurrencyID As Long
    Status As Integer
    TotalVAT As Long
    TotalDiscount As Long
    TotalPayable As Long
    TotalExtension As Long
    TotalExtension_Foreign As Long
    TotalExtensionSimple As Long
    TotalExtensionSimple_Foreign As Long
    TotalVAT_Foreign As Long
    TotalDiscount_Foreign As Long
    TotalPayable_Foreign As Long
    TotalQty As Long
    Memo As String * 250
    RType As String * 1
    DispatchMethodID As Long
    Log As String * 500
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type RData
    buffer As String * 946
End Type

Public Type RLProps
    RLID As Long
    TRID As Long
    QtyAvailable As Long
    QtyRequested As Long
    QtyApproved As Long
    QtyReturned As Long
    QtyCounted As Long
    QtyRejected As Long
    Price As Long
    DELLID As Long
    DELID As Long
    Sequence As Long
    ForeignPrice As Long
    PID As String * 40
    Title As String * 70
    MainAuthor As String * 20
    Discount As Double
    VATRate As Double
    code As String * 20
    CodeF As String * 20
    EAN As String * 13
    SINVRef As String * 200
    SINVDate As Date
    DocRef As String * 20
    DOCDate As Date
    Note As String * 50
    Pubcode As String * 10
    Status As String * 3
    Section As String * 20
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type RLData
    buffer As String * 567
End Type

Public Type dRProps
    TRID As Long
    TPID As Long
    DOCCode As String * 10
    ApprovalRef As String * 20
    ApprovalTermDate As Date
    DOCDate As Date
    CaptureDate As Date
    TPNAME As String * 100
    TPACCNum As String * 14
    Pubcode As String * 10
    RType As String * 1
    StaffID As Long
    Status As Integer
    ReturnStatus As Integer
    QtySystem As Long
    QtyRequested As Long
    QtyApproved As Long
    QtyReturned As Long
    QtyCounted As Long
End Type
Public Type dRData
    buffer As String * 187
End Type

Public Type dRLProps
    TRID As Long
    RLID As Long
    DELLID As Long
    QtySystem As Long
    QtyRequested As Long
    QtyApproved As Long
    QtyReturned As Long
    QtyCounted As Long
    val As Long
    PID As String * 40
    ProductDescription As String * 50
    Section As String * 25
    code As String * 20
    EAN As String * 15
    Publisher As String * 20
    Pubcode As String * 10
    SupplierInvoiceRef As String * 100
    SupplierInvoiceDate As Date
    Note As String * 100
    
End Type
Public Type dRLData
    buffer As String * 314
End Type

Sub sizedR()
Dim x As RProps
MsgBox LenB(x) / 2
End Sub

