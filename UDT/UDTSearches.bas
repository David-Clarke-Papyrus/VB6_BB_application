Attribute VB_Name = "UDTSearches"
Option Explicit

Public Type SearchProps
  PID As String * 40
  code As String * 15
  EAN As String * 20
  CodeF As String * 19
  Title As String * 50
  Author As String * 50
  Status As String * 5
  Price As Double
  QtyOnHand As Long
  QtyonOrder As Long
  QtyOnBackorder As Long
  QtyTotalSold As Long
  UKPrice As String * 10
  USPrice As String * 10
  LocalPrice As String * 10
  LastDateDelivered As Date
  Publisher As String * 50
  PublicationDate As String * 500
  PublicationPlace As String * 500
  Edition As String * 200
  Distributor As String * 100
  DistributorCode As String * 100
  CopyPrice As Long
  PurchaseDate As Date
  ImageFilename As String * 51
  SoldDate As Date
  copiesSold As Long
  Serial As Long
  SalesList As String * 250
  Obsolete As Boolean
  Categories As String * 250
  Multibuy As String * 50
  Img() As Byte
  Length As Double
  Width As Double
  
End Type
Public Type BookSearchData
    buffer As String * 2324
End Type

Public Type ProductSearchProps
  PID As String * 40
  code As String * 15
  Title As String * 50
  Author As String * 50
  Price As Double
  Stock As Long
  Publisher As String * 50
End Type


Sub testsear()
Dim x As SearchProps
    MsgBox LenB(x) & "     " & LenB(x) / 2
End Sub

