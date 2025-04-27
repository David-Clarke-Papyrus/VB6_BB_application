VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arExchangeSalesLines 
   Caption         =   "ActiveReport2"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19920
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35137
   _ExtentY        =   14923
   SectionData     =   "arExchangeSalesLines.dsx":0000
End
Attribute VB_Name = "arExchangeSalesLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cmSales As Collection
Dim i As Long

Public Sub component(SL As Collection)
    Set cmSales = SL
    i = 0
End Sub

Private Sub Detail_Format()
    i = i + 1
    If cmSales.Count >= i Then
        fEAN = cmSales.Item(i).CodeF
        fTitle = cmSales.Item(i).Title
        fQty = cmSales.Item(i).Qty
        fPrice = cmSales.Item(i).PriceF
        fDisc = cmSales.Item(i).DiscountRateF
        fTotal = cmSales.Item(i).ExtensionF

        Detail.PrintSection
    End If

End Sub
