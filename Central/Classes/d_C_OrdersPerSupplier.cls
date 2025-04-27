VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_OrdersPerSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mudtProps As OPSProps

Public Property Get Code() As String
    Code = Trim$(mudtProps.Code)
End Property

Public Property Get dateOfOrder() As String
    dateOfOrder = Format(mudtProps.dateOfOrder, "dd/mm/yyyy")
End Property
Public Property Get dateOfOrderForSort() As String
    dateOfOrderForSort = Format(mudtProps.dateOfOrder, "yyyy/mm/dd")
End Property

Public Property Get Title() As String
    Title = Trim$(mudtProps.Title)
End Property
Public Property Get Price() As String
    Price = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property

Public Property Let Code(Value As String)
    mudtProps.Code = Value
End Property

Public Property Let dateOfOrder(Value As String)
    mudtProps.dateOfOrder = Value
End Property

Public Property Let Title(Value As String)
    mudtProps.Title = Value
End Property
Public Property Let Price(Value As String)
    mudtProps.Price = Value
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As OPSData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub



