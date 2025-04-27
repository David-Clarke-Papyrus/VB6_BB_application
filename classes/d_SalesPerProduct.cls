VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "d_SalesPerProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mudtProps As SPCProps

Public Property Get code() As String
    code = Trim$(mudtProps.code)
End Property
Public Property Get WeekOfYear() As Integer
    WeekOfYear = DatePart("ww", mudtProps.dateOfSale, vbMonday, vbFirstFourDays)
End Property
Public Property Get MonthOfYear() As Integer
    MonthOfYear = DatePart("m", mudtProps.dateOfSale, vbMonday, vbFirstFourDays)
End Property

Public Property Get dateOfSaleF() As String
    dateOfSaleF = Format(mudtProps.dateOfSale, "dd/mm/yyyy")
End Property
Public Property Get dateOfSaleForSort() As String
    dateOfSaleForSort = Format(mudtProps.dateOfSale, "yyyy/mm/dd")
End Property

Public Property Get Title() As String
    Title = Trim$(mudtProps.Title)
End Property
Public Property Get Price() As String
    Price = Format(mudtProps.Price / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get WeekNumber() As Integer
    WeekNumber = mudtProps.Week
End Property

'Public Property Let Valu(Value As Long)
'    udtProps.Valu = Value
'End Property
'
'Public Property Let Qty(Value As Long)
'    udtProps.Qty = Value
'End Property

Public Property Get Valu()
    Valu = mudtProps.Valu
End Property
Public Property Get ValuF() As String
    ValuF = Format(mudtProps.Valu / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
End Property
Public Property Get Qty()
    Qty = mudtProps.Qty
End Property

Private Sub SetState(buffer As String)
  Dim mudtData As SPCData
  
  mudtData.buffer = buffer
  LSet mudtProps = mudtData
End Sub
Friend Sub SetSuperState(buffer As String)
  SetState buffer
End Sub


