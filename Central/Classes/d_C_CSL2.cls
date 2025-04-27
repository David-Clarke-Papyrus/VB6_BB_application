VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_CSL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mudtProps As CSLProps

Public Property Get Qty() As Long
    Qty = FNN(mudtProps.Qty)
End Property

Public Property Get Price() As Long
    Price = FNN(mudtProps.Price)
End Property
Public Property Get PriceF() As String
    PriceF = Format(FNN(mudtProps.Price) / 100, "###,##0.00")
End Property
Public Property Get Extension() As Long
    Extension = mudtProps.Qty * (mudtProps.Price * ((100 - mudtProps.DiscountRate) / 100))
End Property
Public Property Get ExtensionF() As String
    ExtensionF = Format(Extension / 100, "###,##0.00")
End Property

Public Property Get Discount() As Long
    Discount = FNN(mudtProps.Discount)
End Property
Public Property Get DiscountF() As String
    DiscountF = Format(FNN(mudtProps.Discount) / 100, "###,##0.00")
End Property
Public Property Get VATValue() As Long
    VATValue = Extension * ((mudtProps.VATRate) / 100)
End Property
Public Property Get VATValueF() As String
    VATValueF = Format(VATValue - (VATValue * (100 / (VATRate + 100))), "###,##0.00")
End Property
Public Property Get VATRate() As Double
    VATRate = FNDBL(mudtProps.VATRate)
End Property
Public Property Get VATRateF() As String
    VATRateF = PBKSPercentF(VATRate)
End Property
Public Property Get VATCombinedF() As String
Dim str As String
    If mudtProps.VATRate <> oPC.Configuration.VATRate Then
        str = VATValueF & " (" & VATRateF & ")"
        VATCombinedF = str
    Else
        VATCombinedF = VATValueF
    End If
End Property


Public Property Get DiscountRateF() As String
    DiscountRateF = PBKSPercentF(FNDBL(mudtProps.DiscountRate))
End Property
Public Property Get DiscountCombinedF() As String
Dim str As String
   ' str = Format(mudtProps.Discount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
   ' If mudtProps.DiscountRate > 0 Then
   '     str = str & " (" & DiscountRateF & ")"
        DiscountCombinedF = DiscountRateF
   ' End If
End Property

Public Property Get DiscountRate() As Double
    DiscountRate = FNDBL(mudtProps.DiscountRate)
End Property

Public Property Get CodeF() As String
    CodeF = FNS(mudtProps.CodeF)
End Property

Public Property Get Code() As String
    Code = FNS(mudtProps.Code)
End Property

Public Property Get TitleAuthor() As String
    TitleAuthor = FNS(mudtProps.Title) & "(" & FNS(mudtProps.Author) & ")"
End Property

Public Property Get Title() As String
    Title = FNS(mudtProps.Title)
End Property

Public Property Get Author() As String
    Author = FNS(mudtProps.Author)
End Property

Public Property Get TRID() As Long
    TRID = FNN(mudtProps.TRID)
End Property
Public Property Get ExchangeGUID() As String
    ExchangeGUID = FNS(mudtProps.EXCHANGE_GUID)
End Property


Private Sub SetState(buffer As String)
  Dim udtData As CSLData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
