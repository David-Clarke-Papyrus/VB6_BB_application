Attribute VB_Name = "UDTPTProps"
Public Type PTProps
  ID  As Long
  Active As Boolean
  code As String * 50
  Number As String * 10
  Round As Long
  Discount As Double
  B1Min As Long
  B1Max As Long
  B1MU As Long
  B2Min As Long
  B2Max As Long
  B2MU As Long
  B3Min As Long
  B3Max As Long
  B3MU As Long
  CRSALES As String * 20
  CRSALES_CONTRA As String * 20
  CASHSALES As String * 20
  CASHSALES_CONTRA As String * 20
  PURCHASES As String * 20
  PURCHASES_CONTRA As String * 20
  VAT As String * 20
  SaleOrReturn As Boolean
  IsVoucher As Boolean
  dbactionStatus As Integer
  SystemCode As String * 4
  
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
End Type
Public Type PTData
    buffer As String * 238
End Type



Sub testPTProps()
Dim x As PTProps
    MsgBox LenB(x) / 2
End Sub

