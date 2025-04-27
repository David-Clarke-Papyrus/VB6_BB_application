Attribute VB_Name = "UDTAddress"
Public Type AddressProps
    ID As Long
    TPID As Long
    CountryID As Long
    PostageType As Integer
    Description As String * 15
    Line1 As String * 50
    Line2 As String * 50
    Line3 As String * 50
    Line4 As String * 50
    Line5 As String * 50
    Line6 As String * 50
    pCode As String * 10
    Addressee As String * 80
    Delivery As String * 50
    BusPhone As String * 50
    Phone As String * 50
    Fax As String * 50
    EMail As String * 100
    GetsCatalogue As Boolean
    ForMailing As Boolean
    Category As Integer
    Appro As Boolean
    BillTo As Boolean
    DelTo As Boolean
    OrderTo As Boolean
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type
Public Type AddressData
    buffer As String * 722
End Type

Sub lenaddress()
Dim x As AddressProps
    MsgBox LenB(x) & "    " & LenB(x) / 2
End Sub
