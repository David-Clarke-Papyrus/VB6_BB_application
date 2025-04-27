VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_ZExchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim sExchGUID As String * 16
Dim dtSaleDate As Date
Dim sTillCode As String * 15
Dim lCashTotal As Long
Dim lCardTotal As Long
Dim lChequeTotal As Long
Dim lVoucherTotal As Long
Dim iIndex As Integer


Public Property Get ExchangeGUID() As String
    ExchangeGUID = sExchGUID
End Property
Friend Property Let ExchangeGUID(Val As String)
    sExchGUID = Val
End Property

Public Property Get SaleDateTime() As Date
    SaleDateTime = dtSaleDate
End Property
Friend Property Let SaleDateTime(Val As Date)
    dtSaleDate = Val
End Property

Public Property Get TillCode() As String
    TillCode = sTillCode
End Property
Friend Property Let TillCode(Val As String)
    sTillCode = Val
End Property

Public Property Get CashTotal() As Double
    CashTotal = lCashTotal \ 1000
End Property
Friend Property Let CashTotal(Val As Double)
    lCashTotal = CLng(Val * 100)
End Property

Public Property Get CardTotal() As Double
    CardTotal = lCardTotal / 100
End Property
Friend Property Let CardTotal(Val As Double)
    lCardTotal = CLng(Val * 100)
End Property

Public Property Get ChequeTotal() As Double
    ChequeTotal = lChequeTotal / 100
End Property
Friend Property Let ChequeTotal(Val As Double)
    lChequeTotal = CLng(Val * 100)
End Property

Public Property Get VoucherTotal() As Double
    VoucherTotal = lVoucherTotal / 100
End Property
Friend Property Let VoucherTotal(Val As Double)
    lVoucherTotal = CLng(Val * 100)
End Property

Friend Property Let Index(Val As Integer)
    iIndex = Index
End Property
Public Property Get Index() As Integer
    Index = iIndex
End Property

