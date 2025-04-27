VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mudtProps As PaymentProps

Public Property Get Amt() As Long
    Amt = FNN(mudtProps.Amt)
End Property
Public Property Get AmtF() As String
    AmtF = Format(FNN(mudtProps.Amt) / 100, "###,##0.00")
End Property
Public Property Get AmtTendered() As Long
    AmtTendered = FNN(mudtProps.AmtTendered)
End Property
Public Property Get AmtTenderedF() As String
    AmtTenderedF = Format(FNN(mudtProps.AmtTendered) / 100, "###,##0.00")
End Property
Public Property Get Note() As String
    Note = FNS(mudtProps.Note)
End Property
Public Property Get Reference() As String
Dim s As String
    s = FNS(mudtProps.Reference)
'    If PaymentTypeF = "V" Then
'        Select Case (Left(s))
'
'        End Select
'    Else
        Reference = s
'    End If
End Property
 
Public Property Get PaymentType() As String
    PaymentType = FNS(mudtProps.Type)
End Property
Public Property Get PaymentTypeF() As String
    Select Case UCase(Trim(mudtProps.Type))
    Case "C"
        PaymentTypeF = "Cash"
    Case "A"
        PaymentTypeF = "Credit card"
    Case "V"
        PaymentTypeF = "Voucher"
    Case "Q"
        PaymentTypeF = "Cheque"
    Case "RDE"
        PaymentTypeF = "Refund deposit"
    Case "W"
        PaymentTypeF = "Petty: " & Note
    Case "CV"
        PaymentTypeF = "Credit voucher"
    Case "CNR"
        PaymentTypeF = "Change voucher"
    Case "RD"
        PaymentTypeF = "Red. deposit"
    Case "DDP"
        PaymentTypeF = "Dirct. deposit"
    Case "AC"
        PaymentTypeF = "Account"
    Case Else
        PaymentTypeF = "Unknown"
    End Select
End Property

Public Property Get EXCHANGE_GUID() As String
    EXCHANGE_GUID = FNS(mudtProps.EXCHANGE_GUID)
End Property


Public Property Get ID() As Long
    ID = mudtProps.PaymentID
End Property


Private Sub SetState(buffer As String)
  Dim udtData As PaymentData
  
  udtData.buffer = buffer
  LSet mudtProps = udtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub
