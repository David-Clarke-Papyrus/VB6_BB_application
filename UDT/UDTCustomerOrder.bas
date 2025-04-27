Attribute VB_Name = "UDTCOrder"
Public Type COrderProps
    TRID As Long
    CustomerID As Long
    BillToAddressID As Long
    GoodsAddressID As Long
    DiscountPercent As Long
    StaffID As Long
    OrderNum As String * 20
    ForAttn As String * 100
    ETA As Date
    Log As String * 500
    ORGUID As String * 40
    DOCDate As Date
    CaptureDate As Date
    QuotedDeliveryCharge As Currency
    TPID As Long
    StaffName As String * 10
    TPNAME As String * 100
    TPPhone As String * 25
    TPFax As String * 25
    TPACCNum As String * 14
    Memo As String * 350
    BusPhone As String * 25
    DepositPaid As Currency
    DepositRefunded As Currency
    VerificationNum As String * 20
    CardNum As String * 40
    ExpiryDate As String * 20
    PaymentMethod As String * 30
    DOCCode As String * 14
    TotalExtras As Long
    TotalDiscount As Long
    TotalNonVAT As Long
    TotalVAT As Long
    TotalExtension As Long
    TotalPayable As Long
    TotalServiceItem As Long
    TotalQty As Long
    VATable As Boolean
    Status As Integer
    OrderType As Integer '0 - ordinary order, 1 = Want,  2 = Standing order
    ShowVAT As Boolean
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type COrderData
    buffer As String * 1398
End Type

Public Type COLProps
    Sequence As Long
    COLineID As Long
    TRID As Long
    SupplierID As Long
    ProdID As String * 40
    Qty As Long
    QtyFirm As Long
    QtySS As Long
    QtyDispatched As Long
    ActionTaken As Long
    Fulfilled As String * 4
    Discount As Double
    Price As Long
    ForeignPrice As Long
    FCFactor As Double
    FCID As Long
    CSLID As Long
    POLID As Long
    Replacementfor As Long
    INVLineID As Long
    Deposit As Long
    DepositStatus As String * 1
    Ref As String * 50
    code As String * 30
    CodeF As String * 20
    CodeForExport As String * 16
    EAN As String * 13
    Title As String * 120
    Author As String * 30
    WantDate As Date
    ETA As Date
    ETACode As String * 2
    Publisher As String * 20
    ExtraPID As String * 40
    ExtraCode As String * 20
    ExtraCharge As Long
    ExtraChargeDescription As String * 50
    Note As String * 350
    LastAction As String * 170
    DeliveryDocument As String * 20
    LastActionDate As Date
    DateReplaced As Date
    VATRate As Double
    ExtraVATRate As Double
    
    ServiceItem As Boolean
    Status As Integer
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type

Public Type COLData
    buffer As String * 1076
End Type

Public Type COProps
    ID As Long
    TRID As Long
    code As String * 50
    Date As Date
    COLFulFilled As String * 20
    COLQty As Integer
    COLCollected As Integer
    
End Type
Public Type COData
    buffer As String * 80
End Type
Public Type dPaymentProps
  DepositID As Long
  DepositorName As String * 50
  DepositorFullName As String * 100
  DepositorPhone As String * 25
  DepositorAcNo As String * 10
  RemittanceReference As String * 50
  PaymentID As Long
  CustomerName As String * 50
  CustomerFullName As String * 100
  CustomerPhone As String * 25
  CustomerAcno As String * 10
  DepositDate As Date
  CaptureDate As Date
  DepositCode As String * 10
  StaffID As Long
  DepositorID As Long
  DepositAmount As Double
  DepositSettlementDiscount As Double
  DepositType As String * 20
  StatusF As String * 20
  Status As Integer
End Type
Public Type dPaymentData
    buffer As String * 500
End Type


Public Type dCOProps
  TRID As Long
  TPNAME As String * 50
  TPPhone As String * 25
  TPAccNo As String * 10
  EMail As String * 100
  CustomerDisplay As String * 100
  DOCDate As Date
  CaptureDate As Date
  ETA As Date
  DOCCode As String * 10
  StaffID As Long
  COLID As Long
  PID As String * 40
  Qty As Long
  QtyDispatched As Long
  Price As Long
  Deposit As Long
  Discount As Double
  Ref As String * 20
  StatusF As String * 20
  Status As Integer
  Used As Boolean
  OrderType As Integer
End Type
Public Type dCOData
    buffer As String * 412
End Type
Public Type dCOLProps
    COLID As Long
    TRID As Long
    Qty As Long
    QtyDispatched As Long
    QtyToAllocate As Long  'this is used when this record is used in c_COLsperDELL
    QtyOnHand As Long
    QtyFirm As Long
    QtySS As Long
    Price As Long
    Discount As Double
    DOCDate As Date
    ETA As Date
    LastActionDate As Date
    WantDate As Date
    DOCCode As String * 10
    PID As String * 40
    Ref As String * 50
    Title As String * 30
    code As String * 13
    CustName As String * 25
    CustInitials As String * 6
    CustTitle As String * 7
    Note As String * 2350
    Reports As String * 300
    LastAction As String * 150
End Type
Public Type dCOLData
    buffer As String * 3020
End Type

Public Type COLLIBProps
    Qty As Long
    Discount As Double
    Price As Double
    CSLID As Long
    Ref As String
    EAN As String
    Title As String
    Author As String
    Publisher As String
    Edition As String
    PublicationDate As String
    Note As String
    VATRate As Double
    Status As Integer
End Type

Public Type dCOSRProps
    TRID As Long
    TPNAME As String * 100
    CustomerDisplay As String * 100
    DOCCode As String * 10
    DOCDate As Date
    CaptureDate As Date
    TPAccNo As String * 40
    Status As String * 5
    StaffID As Long
    SM As String * 100
    Log As String * 500
End Type
Public Type dCOSRData
    buffer As String * 868
End Type


Sub GetCOLLen()
Dim x As COrderProps
MsgBox LenB(x) / 2
End Sub
