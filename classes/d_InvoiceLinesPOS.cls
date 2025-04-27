VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_InvoiceLinesPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private udtProps As InvoiceLinesProps

Private Sub SetState(buffer As String)
  Dim mudtData As InvoiceLinesData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Public Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
Public Property Get ILID() As Long
    ILID = FNS(udtProps.ILID)
End Property

Public Property Get PID() As String
    PID = FNS(udtProps.PID)
End Property
    
Public Property Get Description() As String
    Description = FNS(udtProps.Description)
End Property
    
Public Property Get LineCode() As String
    LineCode = FNS(udtProps.LineCode)
End Property

Public Property Get DiscountRate() As Double
    DiscountRate = FNDBL(udtProps.DiscountRate)
End Property
Public Property Get DiscountRateF() As String
    DiscountRateF = PBKSPercentF(udtProps.DiscountRate)
End Property

Public Property Get PriceF() As String
    PriceF = Format(FNN(udtProps.Price) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Property
Public Property Get Price() As Double
    Price = FNDBL(udtProps.Price)
End Property
Public Property Get Qty() As Long
    Qty = FND(udtProps.Qty)
End Property
Public Property Get QtyCredited() As Long
    QtyCredited = FND(udtProps.QtyCredited)
End Property

Public Property Get LineExtension() As Double
    LineExtension = (Qty * (Price / oPC.CurrencyDivisor)) * ((100 - DiscountRate) / 100)
End Property

Public Property Get LineExtensionF() As String
    LineExtensionF = Format(LineExtension, oPC.CurrencyFormat)
End Property



