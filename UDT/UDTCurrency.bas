Attribute VB_Name = "UDTCurrency"
Public Type CurrencyProps
  ID  As Long
  Symbol As String * 1
  Format As String * 100
  Divisor As Long
  Factor As Double
  Description As String * 50
  SYS As String * 3
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
End Type
Public Type CurrencyData
    buffer As String * 188
End Type

