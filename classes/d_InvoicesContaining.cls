VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_InvoicesContaining"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private udtProps As InvoicesContainingProps

Private Sub SetState(buffer As String)
  Dim mudtData As InvoicesContainingData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
Public Property Get TRID() As Long
    TRID = FNN(udtProps.TRID)
End Property
Public Property Get DocCode() As String
    DocCode = FNS(udtProps.DocCode)
End Property
Public Property Get PID() As String
    PID = FNS(udtProps.PID)
End Property


Public Property Get Amount() As Long
    Amount = FNN(udtProps.DocValue)
End Property
Public Property Get AmountF() As String
    AmountF = Format(FNN(udtProps.DocValue) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Property
Public Property Get InvoiceDate() As Date
    InvoiceDate = FND(udtProps.DOCDate)
End Property
Public Property Get InvoiceDateF() As String
    InvoiceDateF = Format(FNS(udtProps.DOCDate), "dd/mm/yyyy")
End Property






