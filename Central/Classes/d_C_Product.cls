VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Product"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private udtProps As SearchProps
Public Property Get PID() As String
  PID = Trim$(udtProps.PID)
End Property

Public Property Get Code() As String
  Code = FNS(udtProps.Code)
End Property
Public Property Get CodeF() As String
  CodeF = FNS(udtProps.CodeF)
End Property
Public Property Get DistributorByIdx(pIdx As Integer) As String
  DistributorByIdx = FNS(udtProps.Distributor)
End Property

Public Property Get Title() As String
  Title = FNS(udtProps.Title)
End Property
Public Property Get LocalPrice() As Long
  LocalPrice = udtProps.LocalPrice
End Property
Public Property Get LocalPriceF() As String
    LocalPriceF = FNS(Format(udtProps.LocalPrice / oPC.Configuration.DefaultCurrencyDivisor, oPC.Configuration.DefaultCurrencyFormatString))
End Property
Public Property Get UKPriceF() As String
    UKPriceF = FNS(Format(udtProps.UKPrice / oPC.Configuration.DefaultCurrencyDivisor, "#,###.00"))
End Property
Public Property Get UKPrice() As Long
  UKPrice = udtProps.UKPrice
End Property
Public Property Get USPriceF() As String
    USPriceF = FNS(Format(udtProps.USPrice / oPC.Configuration.DefaultCurrencyDivisor, "#,###.00"))
End Property
Public Property Get USPrice() As Long
  USPrice = udtProps.USPrice
End Property

Public Property Get Author() As String
  Author = FNS(udtProps.Author)
End Property
Public Property Get Status() As String
  Status = FNS(udtProps.Status)
End Property
Public Property Get StatusF() As String
    Select Case FNS(udtProps.Status)
    Case "O"
        StatusF = "Out of print"
    Case "R"
        StatusF = "Awaiting reprint"
    Case "C"
        StatusF = "Cancelled before printing"
    Case Else
        StatusF = ""
    End Select
End Property

Public Property Get Price() As Double
  Price = udtProps.Price
End Property
Public Property Get EAN() As String
  EAN = FNS(udtProps.EAN)
End Property
Public Property Get PubDate() As String
  PubDate = FNS(udtProps.PublicationDate)
End Property
Public Property Get PubPlace() As String
  PubPlace = FNS(udtProps.PublicationPlace)
End Property

Public Property Get QtyTotalSold() As Long
  QtyTotalSold = udtProps.QtyTotalSold
End Property
Public Property Get QtyOnHand() As Long
  QtyOnHand = udtProps.QtyOnHand
End Property
Public Property Get QtyOnOrder() As Long
  QtyOnOrder = udtProps.QtyOnOrder
End Property
Public Property Get QtyOnBackorder() As Long
  QtyOnBackorder = udtProps.QtyOnBackorder
End Property
Public Property Get LastDateDelivered() As Date
  LastDateDelivered = udtProps.LastDateDelivered
End Property
Public Property Get LastDateDeliveredF() As String
  LastDateDeliveredF = Format(udtProps.LastDateDelivered, "dd/mm/yyyy")
End Property

Public Property Get Publisher() As String
  Publisher = FNS(udtProps.Publisher)
End Property

Friend Property Let PID(Value As String)
  udtProps.PID = Value
End Property

Friend Property Let Code(Value As String)
  udtProps.Code = Value
End Property
Friend Property Let CodeF(Value As String)
  udtProps.CodeF = Value
End Property
Friend Property Let Distributor(Value As String)
  udtProps.Distributor = Value
End Property

Friend Property Let Title(Value As String)
  udtProps.Title = Value
End Property
Friend Property Let LocalPrice(Value As Long)
  udtProps.LocalPrice = Value
End Property
Friend Property Let UKPrice(Value As Long)
  udtProps.UKPrice = Value
End Property
Friend Property Let USPrice(Value As Long)
  udtProps.USPrice = Value
End Property

Friend Property Let Author(Value As String)
  udtProps.Author = Value
End Property
Friend Property Let EAN(Value As String)
  udtProps.EAN = Value
End Property
Friend Property Let Status(Value As String)
  udtProps.Status = Value
End Property

Friend Property Let Price(Value As Double)
  udtProps.Price = Value
End Property

Friend Property Let QtyOnHand(Value As Long)
  udtProps.QtyOnHand = Value
End Property
Friend Property Let QtyOnOrder(Value As Long)
  udtProps.QtyOnOrder = Value
End Property
Friend Property Let QtyOnBackorder(Value As Long)
  udtProps.QtyOnBackorder = Value
End Property
Friend Property Let QtyTotalSold(Value As Long)
  udtProps.QtyTotalSold = Value
End Property
Friend Property Let LastDateDelivered(Value As Date)
  udtProps.LastDateDelivered = Value
End Property
'Friend Property Let Stock(Value As Long)
'  udtProps.Stock = Value
'End Property

Friend Property Let Publisher(Value As String)
  udtProps.Publisher = Value
End Property

Public Property Let SalesList(Value As String)
  udtProps.SalesList = Value
End Property
Public Property Get SalesList() As String
  SalesList = FNS(udtProps.SalesList)
End Property

Public Property Get Distributor() As String
  Distributor = FNS(udtProps.Distributor)
End Property


