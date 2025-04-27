Attribute VB_Name = "UDT_Internal"
Option Explicit

Public Type restruct           ' Defining restruct to be a byte array
  Temp(1 To 8192) As Byte
End Type

Public Type tTotal
    Description As String
    val As Long
    AmtFormatted As String
    TotalType As String
    RunningTotal As Long
    Sign As String
End Type

 Public Type OldPrices
    UKPrice As Long
    USPrice As Long
    EUPrice As Long
    RRP As Long
    SpecialPrice As Long
    SP As Long
    Cost As Long
    MultibuyCode As String
    NDA As Boolean
End Type
'
Public Type OldCustomerDiscounts
    Discount As Double
    CreditLimit As Long
    Terms As Long
    Blocked As Boolean
End Type
Public Type TTotal2
    Description As String
    val As Long
    AmtFormatted As String
    TotalType As String
    RunningTotal As Long
    Sign As String
End Type

Public Type SummType
  PID As String
  QtyAlloc As Long
  QtyOnHand As Long
End Type
Public Type Reserve
    Title As String
    Author As String
    Customer As String
    Phone As String
    DateOrdered As String
    DateReceived As String
    Qty As String
End Type
