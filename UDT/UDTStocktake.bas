Attribute VB_Name = "UDTStocktake"
Public Type StockTakeProps
 '   ID As Long
  '  SAID As Long
    TRID As Long
    OperatorID As Long
    Note As String * 255
    Code As String * 10
    NominalDate As Date
    CaptureDate As Date
    CutoffDate As Date
    ValueOfStockRetail As Currency
    ValueOfStockCost As Currency
    TotalProducts As Long
    ItemsWithoutPrices As Long
    ProductsWithoutPrices As Long
    AvgDiscount As Double
    TotalItems As Long
    Status As Integer
    Zeroising As Boolean
    
    Completed As Boolean
    Void As Boolean
    Printed As Boolean
    
    IsNew As Boolean
    IsDeleted As Boolean
    IsDirty As Boolean
End Type

Public Type StocktakeData
     buffer As String * 326
End Type

Sub stsize()
Dim X As StockTakeProps
    MsgBox LenB(X) & "   " & LenB(X) / 2
End Sub
