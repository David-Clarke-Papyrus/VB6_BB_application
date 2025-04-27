Attribute VB_Name = "UDTConfiguration"
Public Type ConfigProps
    UsesBookfind As Boolean
    UsesGardners As Boolean
    SupportsLoyaltyClub As Boolean
    LookupSeq As String * 10
    LocalCountryID As Long
    DefaultCurrencyID As Long
    LocalCurrencyID As Long
    VATRate As Double
    UnallocatedPT As Long
    CustomerTypeDictID As Long
    CustomerIGDictID As Long
    LoyaltyClubTypeID As Long
    LastStockTakeDate As Date
    DefaultStoreID As Long
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type ConfigData
    buffer As String * 42
End Type

Sub configLen()
Dim X As ConfigProps
    MsgBox LenB(X) & "   " & LenB(X) / 2
End Sub
