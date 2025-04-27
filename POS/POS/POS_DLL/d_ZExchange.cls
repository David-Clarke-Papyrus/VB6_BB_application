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
Dim lCashTotal As Long
Dim lCardTotal As Long
Dim lChequeTotal As Long
Dim lVoucherTotal As Long
Dim iIndex As Integer


Public Property Get ExchangeGUID() As String
    ExchangeGUID = sExchGUID
End Property
Friend Property Let ExchangeGUID(val As String)
    sExchGUID = val
End Property

Public Property Get SaleDateTime() As Date
    SaleDateTime = dtSaleDate
End Property
Friend Property Let SaleDateTime(val As Date)
    dtSaleDate = val
End Property

'Public Property Get TillCode() As String
'    TillCode = sTillCode
'End Property
'Friend Property Let TillCode(val As String)
'    sTillCode = val
'End Property

Public Property Get CashTotal() As Double
    On Error GoTo errHandler
    CashTotal = lCashTotal \ 1000
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_ZExchange.CashTotal"
End Property
Friend Property Let CashTotal(val As Double)
    On Error GoTo errHandler
    lCashTotal = CLng(val * 100)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_ZExchange.CashTotal(Val)", val
End Property

Public Property Get CardTotal() As Double
    On Error GoTo errHandler
    CardTotal = lCardTotal / 100
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_ZExchange.CardTotal"
End Property
Friend Property Let CardTotal(val As Double)
    On Error GoTo errHandler
    lCardTotal = CLng(val * 100)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_ZExchange.CardTotal(Val)", val
End Property

Public Property Get ChequeTotal() As Double
    On Error GoTo errHandler
    ChequeTotal = lChequeTotal / 100
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_ZExchange.ChequeTotal"
End Property
Friend Property Let ChequeTotal(val As Double)
    On Error GoTo errHandler
    lChequeTotal = CLng(val * 100)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_ZExchange.ChequeTotal(Val)", val
End Property

Public Property Get VoucherTotal() As Double
    On Error GoTo errHandler
    VoucherTotal = lVoucherTotal / 100
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_ZExchange.VoucherTotal"
End Property
Friend Property Let VoucherTotal(val As Double)
    On Error GoTo errHandler
    lVoucherTotal = CLng(val * 100)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "d_ZExchange.VoucherTotal(Val)", val
End Property

Friend Property Let Index(val As Integer)
    iIndex = val
End Property
Public Property Get Index() As Integer
    Index = iIndex
End Property

