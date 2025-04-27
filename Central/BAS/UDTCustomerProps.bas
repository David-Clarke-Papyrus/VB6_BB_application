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
    CustomerTypeID As Long
    DefaultDiscount As Double
    DefaultDeliveryDays As Long
    Name As String * 100
    Initials As String * 15
    Title As String * 10
    Mobile As String * 20
    AcNo As String * 15
    SoundexName As String * 10
    Phone As String * 50
    SearchPhone As String * 15
    StoreName As String * 20
    StoreID As Long
    IDNUM As String * 15
    NOTE As String * 500
    VATable As Boolean
    GetsCatalogue As Boolean
    CanBeDeleted As Boolean
    CustNotifyBookLaunch As Boolean
    CustNotifyBookSale As Boolean
    CustNotifyBookPromotion As Boolean
    ExcludeFromSales As Boolean
    DateRecordAdded As Date
    DateLastModified As Date
    
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type CustomerData
     buffer As String * 808
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
     buffer As String * 44
End Type

Public Type CustomerPropsDisplay
    ID  As Long
    Initials As String * 15
    Appell As String * 15
    Name As String * 50
    AcNo As String * 12
    CELL As String * 18
    L1 As String * 25
    L2 As String * 25
    L3 As String * 25
    L4 As String * 25
    L5 As String * 25
    L6 As String * 25
    Country As String * 25
    PostCode As String * 8
    Addressee As String * 25
    FullIdentification As String * 80
    Phone As String * 30
    CustomerTypeID As Long
    CustomerTypeDescription As String * 15
    MailingAddress As String * 150
    Email As String * 100
    DefaultAddressID As Long
    GetsCatalogue As Boolean
    SalesValue As Long
    SalesQty As Long
    DOB As String * 20
    NOTE As String * 200
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type CustomerDataDisplay
     buffer As String * 926
End Type
Public Type SPCProps
    Price As Long
    dateOfSale As Date
    Week As Integer
    Qty As Long
    Valu As Long
    Title As String * 120
    Code As String * 15
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
Dim X As CustomerProps
    MsgBox LenB(X) & "    " & LenB(X) / 2
End Sub
