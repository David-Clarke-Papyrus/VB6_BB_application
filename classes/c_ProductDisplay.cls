VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProductDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private udtProps As SearchProps
Public Property Get BookID() As Long
  BookID = udtProps.BookID
End Property

Public Property Get ISBN() As String
  ISBN = Trim$(udtProps.ISBN)
End Property

Public Property Get Title() As String
  Title = Trim$(udtProps.Title)
End Property

Public Property Get Author() As String
  Author = Trim$(udtProps.Author)
End Property

Public Property Get Price() As Double
  Price = udtProps.Price
End Property

Public Property Get Stock() As Long
  Stock = udtProps.Stock
End Property

Public Property Get Publisher() As String
  Publisher = udtProps.Publisher
End Property

Friend Property Let BookID(Value As Long)
  udtProps.BookID = Value
End Property

Friend Property Let ISBN(Value As String)
  udtProps.ISBN = Value
End Property

Friend Property Let Title(Value As String)
  udtProps.Title = Value
End Property

Friend Property Let Author(Value As String)
  udtProps.Author = Value
End Property

Friend Property Let Price(Value As Double)
  udtProps.Price = Value
End Property

Friend Property Let Stock(Value As Long)
  udtProps.Stock = Value
End Property

Friend Property Let Publisher(Value As String)
  udtProps.Publisher = Value
End Property




