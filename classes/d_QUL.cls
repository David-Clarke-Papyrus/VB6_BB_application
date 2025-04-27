VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_QUL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type DisplayProps
    QUID As Long
    QULID As Long
    code As String
    BookTitle As String
    MainAuthor As String
    Publisher As String
    InvLineMemo As String
    Deposit As Currency
    Discount As Double
    Qty As Integer
    Price As Currency
End Type

Private udtProps As DisplayProps

Public Property Get QUID() As Long
    QUID = udtProps.QUID
End Property
Public Property Let QUID(val As Long)
    udtProps.QUID = val
End Property

Public Property Get QULID() As Long
    QULID = udtProps.QULID
End Property
Public Property Let QULID(val As Long)
    udtProps.QULID = val
End Property

Public Property Get code() As String
    code = udtProps.code
End Property
Public Property Let code(val As String)
    udtProps.code = val
End Property

Public Property Get BookTitle() As String
    BookTitle = udtProps.BookTitle
End Property
Public Property Let BookTitle(val As String)
    udtProps.BookTitle = val
End Property

Public Property Get MainAuthor() As String
    MainAuthor = udtProps.MainAuthor
End Property
Public Property Let MainAuthor(val As String)
    udtProps.MainAuthor = val
End Property

Public Property Get Publisher() As String
    Publisher = udtProps.Publisher
End Property
Public Property Let Publisher(val As String)
    udtProps.Publisher = val
End Property

Public Property Get InvLineMemo() As String
    InvLineMemo = udtProps.InvLineMemo
End Property
Public Property Let InvLineMemo(val As String)
    udtProps.InvLineMemo = val
End Property

Public Property Get Deposit() As Currency
    Deposit = udtProps.Deposit
End Property
Public Property Let Deposit(val As Currency)
    udtProps.Deposit = val
End Property

Public Property Get Discount() As Double
    Discount = udtProps.Discount
End Property
Public Property Let Discount(val As Double)
    udtProps.Discount = val
End Property

Public Property Get Qty() As Integer
    Qty = udtProps.Qty
End Property
Public Property Let Qty(val As Integer)
    udtProps.Qty = val
End Property

Public Property Get Price() As Currency
    Price = udtProps.Price
End Property
Public Property Let Price(val As Currency)
    udtProps.Price = val
End Property
    

