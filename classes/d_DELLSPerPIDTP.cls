VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_DELLSPerPIDTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim oCurr As a_Currency
Dim mudtProps As dDELLProps
 
Public Property Get Status() As String
    Status = mudtProps.Status
End Property

Public Property Let Status(Value As String)
    mudtProps.Status = Value
End Property
Public Property Get StatusF() As String
    If mudtProps.Status = 1 Then
        StatusF = "VOID"
    ElseIf mudtProps.Status = 2 Then
        StatusF = "IN PROCESS"
    ElseIf mudtProps.Status = 3 Then
        StatusF = "ISSUED"
    ElseIf mudtProps.Status = 4 Then
        StatusF = "COMPLETE"
    ElseIf mudtProps.Status = 7 Then
        StatusF = "PROFORMA"
    ElseIf mudtProps.Status = 6 Then
        StatusF = "CANCELLED"
    End If
End Property

Public Property Get SINVDate() As String
    SINVDate = FNS(mudtProps.SINVDate)
End Property


Public Property Get SINVRef() As String
    SINVRef = FNS(mudtProps.SINVRef)
End Property

Property Get DocRef() As String
    DocRef = FNS(mudtProps.DocRef)
End Property

Property Get DOCDate() As String
    DOCDate = Format(mudtProps.DOCDate, "dd/mm/yyyy")
End Property

Property Get Qty() As String
    Qty = mudtProps.Qty
End Property

Public Property Let TRID(Value As Long)
    mudtProps.TRID = Value
End Property
Public Property Get TRID() As Long
    TRID = mudtProps.TRID
End Property
Public Property Get PriceF() As String
    If mudtProps.CURRID <> oPC.Configuration.DefaultCurrencyID Then
        PriceF = Format(mudtProps.ForeignPrice / mudtProps.CurrDivisor, mudtProps.CurrFormatString)
    Else
        PriceF = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    End If
End Property
Public Property Get LocalPrice() As Long
    LocalPrice = mudtProps.Price
End Property
Public Property Get ForeignPrice() As Long
    ForeignPrice = mudtProps.ForeignPrice
End Property


Private Sub SetState(buffer As String)
  Dim udtData As dDELLData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
Dim oC As a_Currency

  SetState buffer
  For Each oC In oPC.Configuration.Currencies
    If oC.ID = mudtProps.CURRID Then
        Set oCurr = oC
        Exit For
    End If
  Next
End Sub

Public Property Get DELLID() As Long
    DELLID = mudtProps.DELLID
End Property
Public Property Get DiscountF() As String
    DiscountF = PBKSPercentF(Round(mudtProps.Discount, 2))
End Property
Public Property Get Discount() As Double
    Discount = Round(mudtProps.Discount, 2)
End Property

