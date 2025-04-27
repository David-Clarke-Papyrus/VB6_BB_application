Attribute VB_Name = "Module1"
Public Type ReturnRec
    Code  As String
    title As String
    PID As String
    APPLID As Long
    APPLQtySold As Long
    APPLQtyReturned As Long
    Price As String
    DiscountRate As Double
    VATRate As Double
End Type

Public Type InvoiceRec
    Code  As String
    title As String
    PID As String
    ILID As Long
    Qty As Long
    QtyCredited As Long
    Price As String
    DiscountRate As Double
    VATRate As Double
End Type

Public Type Reserve
    title As String
    Author As String
    Customer As String
    Phone As String
    DateOrdered As String
    DateReceived As String
    Qty As String
End Type

Public Type OldPrices
    UKPrice As Long
    USPrice As Long
    RRP As Long
    SpecialPrice As Long
    SP As Long
    Cost As Long
End Type

Public Type OldCustomerDiscounts
    Discount As Long
    CreditLimit As Long
    Terms As Long
    Blocked As Boolean
End Type

