VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_APPLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private udtProps As APPLPOSProps

Private Sub SetState(buffer As String)
  Dim mudtData As APPLPOSData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
    
Public Property Get title() As String
    title = FNS(udtProps.title)
End Property
    
'Public Property Get EAN() As String
'    EAN = FNS(udtProps.CodeF)
'End Property
Public Property Get PID() As String
    PID = FNS(udtProps.PID)
End Property

Public Property Get Code() As String
    Code = FNS(udtProps.Code)
End Property
Public Property Get CodeF() As String
    CodeF = FNS(udtProps.CodeF)
End Property

Public Property Get APPLID() As Long
    APPLID = udtProps.APPLID
End Property


Public Property Get QtyOut() As Long
    QtyOut = FNN(udtProps.QtyOut)
End Property

Public Property Get QtyBack() As Long
    QtyBack = FNN(udtProps.QtyBack)
End Property

Public Property Get QtyNett() As Long
    QtyNett = QtyOut - QtyBack
End Property

Public Property Get DiscountRate() As Double
    DiscountRate = FNDBL(udtProps.DiscountRate)
End Property

Public Property Get DiscountRateF() As String
    DiscountRateF = PBKSPercentF(DiscountRate)
End Property

Public Property Get Price() As Long
    Price = FNN(udtProps.Price)
End Property

Public Property Get VATRate() As Double
    VATRate = FNDBL(udtProps.VATRate)
End Property

Public Property Get VATRateF() As String
    VATRateF = PBKSPercentF(VATRate)
End Property

Public Property Get PriceF() As String
    PriceF = Format(CDbl(udtProps.Price / oPC.CurrencyDivisor), oPC.CurrencyFormat)
End Property

Public Property Get Total() As Long
    Total = FNN(udtProps.Price)
End Property






