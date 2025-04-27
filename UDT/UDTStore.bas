Attribute VB_Name = "UDTStore"
Public Type StorePProps
  STID As Long
  QtyOnHand  As Long
  QtyCopiesOnHand  As Long
  QtyonOrder As Long
  QtyOnBackorder As Long
  LastOrderedDate As Date
  LastQtyFirmOrdered As Long
  LastQtySSOrdered As Long
  QtyReserved As Long
  
  SP As Double
  LastDeliveredDate As Date
  LastDeliveredPrice As Double
  LastDeliveredQty As Long
  
  TotalSold As Long
  LastSoldDate As Date
  
  DateLastStocktake As Date
  QtyLastStocktake As Long
  FirstReceivedDate As Date
  
  PID  As String * 40
  StoreName As String * 40
  StoreCode As String * 10
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
End Type
Public Type StorePData
    buffer As String * 138
End Type
Public Type StoreProps
  STID As Long
  Description As String * 50
  BillAddress As String * 250
  DelAddress As String * 250
  VPNAddress As String * 50
  SystemName As String * 10
  code As String * 3
  IsActive As Boolean
  IsExternal As Boolean
  IsNew As Boolean
  IsDirty As Boolean
  IsDeleted As Boolean
End Type
Public Type StoreData
    buffer As String * 620
End Type


Sub TestStoreprops()
Dim x As StorePProps
    MsgBox LenB(x) & "     " & LenB(x) / 2
End Sub
