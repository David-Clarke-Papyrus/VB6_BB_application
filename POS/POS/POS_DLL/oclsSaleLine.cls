VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaleLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim sCode As String * 13
Dim ID As Long
Dim lSellPrice As Long
Dim lNomPrice As Long
Dim siDiscount As Single
Dim lDepCredit As Long
Dim lCostValue As Long
Dim sGUID As String * 16
Dim sDescription As String * 50
Dim sExcCode As String * 1

Public Property Let Code(Val As String)
    If Len(Val) > Len(sCode) Then Err.Raise 426, , "Code to long"
    sCode = Val
End Property
Public Property Get Code() As String
    Code = sCode
End Property


Public Property Get GUID() As String
    GUID = sGUID
End Property

Public Property Let SellPrice(Value As String)
    lSellPrice = Val(Value) * 100
End Property
Public Property Get SellPrice() As String
    SellPrice = Format(lSellPrice / 100, "0.00")
End Property

Public Property Let NominalPrice(Value As String)
    lNomPrice = Val(Value) * 100
End Property
Public Property Get NominalPrice() As String
    NominalPrice = Format(lNomPrice / 100, "0.00")
End Property

Public Property Let CostValue(Value As String)
    lCostValue = Val(Value) * 100
End Property
Public Property Get CostValue() As String
    CostValue = Format(lCostValue / 100, "0.00")
End Property

Public Property Let DepositCredited(Value As String)
    lDepCredit = Val(Value) * 100
End Property
Public Property Get DepositCredited() As String
    DepositCredited = Format(lDepCredit / 100, "0.00")
End Property

Public Property Let Description(Val As String)
    If Len(Val) > Len(sDescription) Then Err.Raise 426, , "Description to long"
    sDescription = Val
End Property
Public Property Get Description() As String
    Description = sDescription
End Property

Public Property Let ExceptionCode(Val As String)
    If Len(Val) > Len(sExcCode) Then Err.Raise 426, , "ExceptionCode to long"
    sExcCode = Val
End Property
Public Property Get ExceptionCode() As String
    ExceptionCode = sExcCode
End Property

Public Property Let Discount(Val As Single)
    siDiscount = Val
End Property
Public Property Get Discount() As Single
    Discount = siDiscount
End Property


