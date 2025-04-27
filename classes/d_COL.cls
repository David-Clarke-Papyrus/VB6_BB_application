VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_COL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private udtProps As COLPOSProps

Private Sub SetState(buffer As String)
  Dim mudtData As COLPOSData
  
  mudtData.buffer = buffer
  LSet udtProps = mudtData
End Sub

Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
    
Public Property Get PID() As String
    PID = FNS(udtProps.PID)
End Property
    
Public Property Get COLID() As Long
    COLID = FNN(udtProps.COLID)
End Property
'Public Property Get DocCode() As String
'    DocCode = FNS(udtProps.DocCode)
'End Property
    
Public Property Get Code() As String
    Code = FNS(udtProps.Code)
End Property
Public Property Get DepositStatus() As String
    DepositStatus = FNS(udtProps.DepositStatus)
End Property
Public Property Get DepositStatusF() As String
    Select Case FNS(udtProps.DepositStatus)
    Case "O"
        DepositStatusF = "Outstanding"
    Case "P"
        DepositStatusF = "Paid to us"
    Case "E"
        DepositStatusF = "Redeemed"
    Case "X"
        DepositStatusF = "Refunded"
    End Select
End Property
    
Public Property Get Deposit() As Long
    Deposit = FNN(udtProps.Deposit)
End Property
Public Property Get DepositF() As String
    DepositF = Format(FNN(udtProps.Deposit) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Property
    
Public Property Get Qty() As Long
    Qty = FNN(udtProps.Qty)
End Property

Public Property Get QtyDispatched() As Long
    QtyDispatched = FNN(udtProps.QtyDispatched)
End Property

Public Property Get Price() As Long
    Price = FNN(udtProps.Price)
End Property
Public Property Get PriceF() As String
    PriceF = Format(FNN(udtProps.Price) / oPC.CurrencyDivisor, oPC.CurrencyFormat)
End Property

Public Property Get Description() As String
    Description = FNS(udtProps.Description)
End Property

Public Property Get DiscountRate() As Double
    DiscountRate = FNDBL(udtProps.DiscountRate)
End Property
Public Property Get DiscountRateF() As String
    DiscountRateF = PBKSPercentF(FNDBL(udtProps.DiscountRate))
End Property

Public Property Get COLDate() As Date
    COLDate = FND(udtProps.COLDate)
End Property
Public Property Get COLDateF() As String
    COLDateF = Format(FNS(udtProps.COLDate), "dd/mm/yyyy")
End Property

Public Property Get COLDateForSORT() As String
    COLDateForSORT = Format(FNS(udtProps.COLDate), "yyyymmdd")
End Property



