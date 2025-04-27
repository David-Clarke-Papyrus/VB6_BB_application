Attribute VB_Name = "UDTProduct"
Public Type ProductProps
    ID As String * 40
    code As String * 20
    CodeF As String * 20
    CodeForExport As String * 20
    EAN As String * 13
    Availability As String * 5
    CategoryID As Long
    ProductTypeID As Long
    LastSupplierName As String * 25
    LastApproto As String * 500
    LastDealDescription As String * 15
    BFClassification As String * 10
    BFDistributorCode As String * 7
    UKPrice As Long
    USPrice As Long
    EUPrice As Long
    RRP As Long
    SP As Long
    SpecialPrice As Long
    SSP As Long
    Cost As Long
    CostLastStockTake As Double
    ForeignOrderedPrice As Long
    ForeignOrderedCURRID As Long
    SupplierID As Long
    PublisherID As Long
    StockBalance As Long
    DealID As Long
    Seesafe As Long
    QtyLastOrdered As Long
    QtyLastDelivered As Long
    QtyLastCounted As Long
    QtylastSold As Long
    QtyTotalSold As Long
    PriceLastCounted As Long
    PriceLastDelivered As Long
    PriceLastOrdered As Long
    PricelastSold As Long
    CatalogueheadingID As Long
    Category As Long
    StckAgeQty6Mnths As Integer
    StckAgeQty12Mnths As Integer
    StckAgeQty18Mnths As Integer
    StckAgeQty18MnthsPlus As Integer
    DefaultDeliveryDays As Integer
    ReturnAvailability As Integer
    LoyaltyRate As Integer
    StckAgeDate As Date
    VATRate As Double
    QtyCopiesOnHand As Long
    QtyonOrder As Long
    QtyOnBackorder As Long
    QtyOnHand As Long
    QtyReserved As Long
    QtyExpectedBack As Long
    Discount As Double
    LastCopySerial As Integer
    Title As String * 900
    SubTitle As String * 3000
    Article As String * 5
    BindingCode As String * 10
    SeriesTitle As String * 300
    FlagText As String * 140
    Author As String * 900
    Publisher As String * 500
    Note As String * 1500
    Comment As String * 1500
    Description As String * 1500
    Summary As String * 1500
    PublicationDate As String * 500
    PublicationPlace As String * 500
    MainSupplierName As String * 30
    Edition As String * 500
    BIC As String * 51
    BICDescription As String * 100
    Binding As String * 7
    Section As String * 25
    Status As String * 20
    DateAdded As Date
    DateLastModified As Date
    DateLastDelivered As Date
    DateLastCounted As Date
    DateLastOrdered As Date
    DateLastSold As Date
    KeepsCopies As Boolean
    SpecialVat As Boolean
    NDA As Boolean
    ProductType As String * 1
    CatalogueHeading As String * 100
    Obsolete As Boolean
    SkipBFWash As Boolean
    PTRound As Long
    PTB1MIN As Long
    PTB1Max As Long
    PTB1MU As Long
    PTB2MIN As Long
    PTB2Max As Long
    PTB2MU As Long
    PTB3MIN As Long
    PTB3Max As Long
    PTB3MU As Long
    PTDiscount As Double
    PTSeesafe As Boolean
    ExcludeFromSales As Boolean
    SupplierConversionToLocalFactor As Double
    SupplierCurrencyID As Long
    MasterCategory As Long
    Weight As String * 20
    Length As Double
    Width As Double
    Multibuy As String * 3
    LastCatCheckCode As String * 10
    LastCatCheckDate As Date
    Core As Boolean
    SPSetBy As Long
    SPSetDate As Date
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type ProductData
     buffer As String * 14482
End Type

Public Type CopyProps
    ID As Long
    PID As String * 40
    Serial As Integer
    PurchaseDate As Date
    SoldDate As Date
    Comment As String * 2000
    Description As String * 2000
    FlagText As String * 500
    SoldTo As String * 25
    Cost As Long
    Price As Long
    LocalPrice As Long
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type CopyData
     buffer As String * 4588
End Type

Public Type CATALPIProps
    ID As Long
    CATALID As Long
    PIID As Long
    Price As Long
    Serial As Integer
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type CATALPIData
     buffer As String * 12
End Type

Public Type CATALPProps
    ID As Long
    CATALID As Long
    PID As String * 45
    Price As Long
    Serial As Integer
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type
Public Type CATALPData
     buffer As String * 56
End Type



Public Type BICProps
    ID As Long
    code As String * 12
    Description As String * 100
    Level As Integer
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type BICData
     buffer As String * 250
End Type



Public Type ProductDisplayProps
    ID As String * 40
    code As String * 13
    Description As String * 255
    Author As String * 50
End Type
Public Type ProductDisplayData
    buffer As String * 358
End Type

Public Type ProductSectionProps
    SECID As Long
    PID As String * 40
    Description As String * 40
    typ As String * 5
    Priority As Long
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type ProductSectionData
     buffer As String * 93
End Type

Public Type ProductCategoryProps
    CatID As Long
    PID As String * 40
    Description As String * 40
    CatValueID As Long
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type ProductCategoryData
     buffer As String * 88
End Type

Sub test2()
Dim x As ProductProps

    MsgBox LenB(x) & "     " & LenB(x) / 2
End Sub
