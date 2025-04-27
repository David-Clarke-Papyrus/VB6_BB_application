Attribute VB_Name = "UDTPOrder"
Option Explicit

Public Type POProps
    TRID As Long
    TPID As Long
    COMPID As Long
    DocStoreID As Long
    DELTOStoreID As Long
    DispatchMethodID As Long
    DOCCode As String * 10
    DOCDate As Date
    CaptureDate As Date
    ProcessingDate As Date
    TPNAME As String * 100
    TPACCNum As String * 14
    Memo As String * 200
    Log As String * 500
    TPPhone As String * 15
    TPPhone2 As String * 25
    TPFax As String * 15
    DelToAddress As String * 150
    BillTOAddress As String * 150
    DelToAddressID As Long
    BillToAddressID As Long
    OrderType As String * 2
    StaffID As Long
    StaffName As String * 10
    Signature As String * 50
    StaffEmail As String * 50
    Status As Integer
    CurrencyID As Long
    CurrencyRate As Double
    TotalExtension As Long
    TotalVAT As Long
    TotalQty As Long
    TotalDiscount As Long
    TotalPayable As Long
    TotalExtensionSimple As Long
    TotalExtensionSimple_Foreign As Long
    TotalExtension_Foreign As Long
    TotalVAT_Foreign As Long
    TotalDiscount_Foreign As Long
    TotalPayable_Foreign As Long
    ContainsCO As Boolean
    TempETA As Date
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type POData
    buffer As String * 1370
End Type

Public Type POLProps
    Sequence As Long
    POLID As Long
    TRID As Long
    QtyFirm As Long
    QtySS As Long
    QtyReceivedSoFar As Long
    Fulfilled As String * 4
    Price As Long
    ForeignPrice As Long
    ProductTypeID As Long
    DealID As Long
    COLID As Long
    PID As String * 40
    Section As String * 51
    Title As String * 255
    MainAuthor As String * 100
    Publisher As String * 100
    PublicationDate As String * 35
    Edition As String * 100
    Discount As Double
    VATRate As Double
    Note As String * 100
    Ref As String * 30
    ETA As Date
    LastActionDate As Date
    LastAction As String * 170
    ETACode As String * 3
    ProductCodeForExport As String * 20
    ProductCodeF As String * 20
    ProductCode As String * 15
    EAN As String * 13
    COQty As Long
    LastActionCode As String * 1
    DateReplaced As Date
    Replacementfor As Long
  
  IsNew As Boolean
  IsDeleted As Boolean
  IsDirty As Boolean
End Type
Public Type POLData
    buffer As String * 1110
End Type
Public Type dPOLProps

    POLID As Long
    TRID As Long
    DOCDate As Date
    Price As Long
    ForeignPrice As Long
    CDivisor As Long
    CForeignDivisor As Long
    QtySS As Long
    QtyFirm As Long
    QtyOutstanding As Long
    QtyFirmOutstanding As Long
    QtySSOutstanding As Long
    CFactor As Double
    QtyReceivedSoFar As Long
    QtyUnMatchedTmp As Long
    QtySSUnMatchedTmp As Long
    QtyFIRMUnMatchedTmp As Long
    SupplierID As Long
    VATRate As Long
    COLID As Long
    Discount As Double
    ETA As Date
    LastActionDate As Date
    Replacementfor As Long
    Ref As String * 50
    DOCCode As String * 15
    SOLFulfilled As String * 20
    CFormat As String * 15
    PID As String * 40
    Title As String * 25
    MainAuthor As String * 25
    code As String * 15
    CodeF As String * 20
    EAN As String * 13
    Supplier As String * 25
    Actions As String * 500
    Note As String * 100
    LastAction As String * 150
    Fulfilled As String * 3
    
    
End Type
Public Type dPOLData
    buffer As String * 1076
End Type

Public Type dPOProps
    TRID As Long
    ETA As Date
    OrderType As String * 6
    CURRID As Long
    DOCCode As String * 20
    DOCDate As Date
    TPNAME As String * 255
    TPAccNo As String * 10
    DispatchMode As String * 30
    Log As String * 300
    QtySS As Long
    QtyFirm As Long
    QtyReceived As Long
    PIID As Long
    POLID As Long
    StaffID As Long
    Discount As Double
    VATRate As Double
    Price As Long
    Title As String * 30
    Status As Integer
    CaptureDate As Date
    POQty   As Long
    POValue As Double
    Parent  As Long
End Type
Public Type dPOData
    buffer As String * 702
End Type


Sub sizedPOL()
Dim x As POData
MsgBox LenB(x) / 2
End Sub
