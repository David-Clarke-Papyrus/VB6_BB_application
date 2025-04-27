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
Attribute VB_Exposed = False
Option Explicit

Private udtProps As SearchProps
Public Property Get PID() As String
  PID = Trim$(udtProps.PID)
End Property
Friend Property Let EAN(Value As String)
  udtProps.EAN = Value
End Property
Public Property Get EAN() As String
  EAN = FNS(udtProps.EAN)
End Property

Public Property Get code() As String
  code = FNS(udtProps.code)
End Property
Public Property Get CodeF() As String
  CodeF = FNS(udtProps.CodeF)
End Property
Public Property Get DistributorByIdx(pIdx As Integer) As String
  DistributorByIdx = FNS(udtProps.Distributor)
End Property
Public Property Get CopyPriceF() As String
    CopyPriceF = FNS(Format(udtProps.CopyPrice / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString))
End Property
Public Property Get PurchaseDateF() As String
    If udtProps.PurchaseDate <> CDate(0) Then
        PurchaseDateF = Format(udtProps.PurchaseDate, "dd/mm/yyyy")
    Else
        PurchaseDateF = ""
    End If
End Property
Public Property Get SoldDateF() As String
    If udtProps.PurchaseDate <> CDate(0) Then
        SoldDateF = Format(udtProps.SoldDate, "dd/mm/yyyy")
    Else
        SoldDateF = ""
    End If
End Property
Public Property Get SerialF() As String
    SerialF = CStr(FNN(udtProps.Serial))
End Property

Public Property Get Title() As String
  Title = FNS(udtProps.Title)
End Property
Public Property Get LocalPrice() As Long
  LocalPrice = udtProps.LocalPrice
End Property
Public Property Get LocalPriceF() As String
    LocalPriceF = FNS(Format(udtProps.LocalPrice / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString))
End Property
Public Property Get UKPriceF() As String
    UKPriceF = FNS(Format(udtProps.UKPrice / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.Currencies.FindBySysname("GBP").FormatString))
End Property
Public Property Get UKPrice() As Long
  UKPrice = udtProps.UKPrice
End Property
Public Property Get USPriceF() As String
    USPriceF = FNS(Format(udtProps.USPrice / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.Currencies.FindBySysname("USD").FormatString))
End Property
Public Property Get USPrice() As Long
  USPrice = udtProps.USPrice
End Property

Public Property Get Author() As String
  Author = FNS(udtProps.Author)
End Property
Friend Property Let Status(Value As String)
  udtProps.Status = Value
End Property
Public Property Get Status() As String
    Status = FNS(udtProps.Status)
End Property

Public Property Get STATUSEx(Optional bShort As Boolean = False, Optional Markers As Boolean) As String
Dim s As String
    s = FNS(udtProps.Status)
    If bShort Then
        If s = "IP" Then s = ""
    End If
    If Markers Then
        If s > "" Then s = "**" & s & "**"
    End If
    STATUSEx = s
End Property

Public Property Get StatusF(Optional bShort As Boolean = False, Optional Markers As Boolean) As String
Dim s As String
    
    If FNB(udtProps.Obsolete) = True Then
        StatusF = "OBSOLETE"
'    Else
'        s = oPC.Configuration.ProductStatus.Item(oPC.Configuration.ProductStatus.IndexByF3(STATUSEx(True, False)))
    End If
    If Markers Then
        If s > "" Then s = "**" & s & "**"
    End If
    StatusF = s
End Property
Public Property Get StatusShortF(Optional bShort As Boolean = False, Optional Markers As Boolean) As String
Dim s As String
    
    If FNB(udtProps.Obsolete) = True Then
        StatusShortF = "OBSOLETE"
    Else
        s = oPC.Configuration.ProductStatus.ItemByF3(STATUSEx(True, False))
    End If
    If Markers Then
        If s > "" Then s = "**" & s & "**"
    End If
    StatusShortF = s
End Property

Public Property Get Price() As Double
  Price = udtProps.Price
End Property
Public Property Get copiesSold() As Long
  copiesSold = udtProps.copiesSold
End Property

Public Property Get QtyTotalSold() As Long
  QtyTotalSold = udtProps.QtyTotalSold
End Property
Public Property Get QtyOnHand() As Long
  QtyOnHand = udtProps.QtyOnHand
End Property
Public Property Get QtyonOrder() As Long
  QtyonOrder = udtProps.QtyonOrder
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
Public Property Get PubDate() As String
  PubDate = FNS(udtProps.PublicationDate)
End Property
Public Property Let PubDate(val As String)
  udtProps.PublicationDate = val
End Property
Public Property Get PubPlace() As String
  PubPlace = FNS(udtProps.PublicationPlace)
End Property
Public Property Let PubPlace(val As String)
  udtProps.PublicationPlace = val
End Property

Public Property Get Distributor() As String
  Distributor = Trim$(udtProps.Distributor)
End Property
Public Property Let Distributor(val As String)
  udtProps.Distributor = val
End Property
Public Property Get DistributorCode() As String
  DistributorCode = Trim$(udtProps.DistributorCode)
End Property
Public Property Let DistributorCode(val As String)
  udtProps.DistributorCode = val
End Property

Friend Property Let PID(Value As String)
  udtProps.PID = Value
End Property
Public Property Get Obsolete() As Boolean
  Obsolete = FNB(udtProps.Obsolete)
End Property
Public Property Let Obsolete(val As Boolean)
  udtProps.Obsolete = val
End Property

Public Property Get Multibuy() As String
  Multibuy = FNS(udtProps.Multibuy)
End Property
Public Property Let Multibuy(val As String)
  udtProps.Multibuy = val
End Property

Public Property Get Categories() As String
  Categories = FNS(udtProps.Categories)
End Property
Public Property Let Categories(val As String)
  udtProps.Categories = val
End Property


Friend Property Let code(Value As String)
  udtProps.code = Value
End Property
Friend Property Let CodeF(Value As String)
  udtProps.CodeF = Value
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

Friend Property Let Price(Value As Double)
  udtProps.Price = Value
End Property

Friend Property Let QtyOnHand(Value As Long)
  udtProps.QtyOnHand = Value
End Property
Friend Property Let QtyonOrder(Value As Long)
  udtProps.QtyonOrder = Value
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
Public Property Get Publisher() As String
  Publisher = Trim$(udtProps.Publisher)
End Property
Friend Property Let Edition(Value As String)
  udtProps.Edition = Value
End Property
Public Property Get Edition() As String
  Edition = FNS(udtProps.Edition)
End Property

Friend Property Let Serial(val As Long)
    udtProps.Serial = val
End Property
Friend Property Let CopyPrice(val As Long)
    udtProps.CopyPrice = val
End Property
Friend Property Let PurchaseDate(val As Date)
    udtProps.PurchaseDate = val
End Property
Friend Property Let SoldDate(val As Long)
    udtProps.SoldDate = val
End Property

Friend Property Let copiesSold(val As Long)
    udtProps.copiesSold = val
End Property

Friend Property Let Length(Value As Double)
  udtProps.Length = Value
End Property
Friend Property Let Width(Value As Double)
  udtProps.Width = Value
End Property
Friend Property Let Img(Value() As Byte)
  udtProps.Img = Value
End Property
Public Property Get Img() As Byte()
  Img = udtProps.Img
End Property

Public Property Get Length() As Double
  Length = udtProps.Length
End Property
Public Property Get Width() As Double
  Width = udtProps.Width
End Property

Public Property Get ImageFilename() As String
  ImageFilename = FNS(udtProps.ImageFilename)
End Property
Public Property Let ImageFilename(val As String)
  udtProps.ImageFilename = FNS(val)
End Property

Public Property Get LengthF() As String
  LengthF = CStr(CDbl(udtProps.Length) / 1000)
End Property

Public Property Get LengthandWidth() As String
   LengthandWidth = CStr(CDbl(udtProps.Length) / 1000) & " x " & CStr(CDbl(udtProps.Width) / 1000)
End Property
Public Property Get WidthF() As String
  WidthF = CStr(CDbl(udtProps.Width) / 1000)
End Property

