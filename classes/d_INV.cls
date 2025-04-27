VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_INV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private udtProps As INVPOSProps

Private Sub SetState(buffer As String)
  Dim mudtData As InvPOSData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
Public Property Get TPID() As Long
    TPID = FNN(udtProps.TPID)
End Property
    
    
Public Property Get INVID() As Long
    INVID = FNN(udtProps.INVID)
End Property
    
Public Property Get DocCode() As String
    DocCode = FNS(udtProps.DocCode)
End Property
Public Property Get CustomerName() As String
    CustomerName = FNS(udtProps.CustomerName)
End Property
    
Public Property Get VAT() As Long
    VAT = FNN(udtProps.VAT)
End Property

Public Property Get Payable() As Long
    Payable = FNN(udtProps.Payable)
End Property
Public Property Get PayableF() As String
    PayableF = Format(FNN(udtProps.Payable) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Property


Public Property Get DOCDate() As Date
    DOCDate = FND(udtProps.DOCDate)
End Property
Public Property Get DOCDateF() As String
    DOCDateF = Format(FNS(udtProps.DOCDate), "dd/mm/yyyy")
End Property






