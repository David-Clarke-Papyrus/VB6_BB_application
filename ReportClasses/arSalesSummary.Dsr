VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesSummary 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   MDIChild        =   -1  'True
   _ExtentX        =   24844
   _ExtentY        =   14393
   SectionData     =   "arSalesSummary.dsx":0000
End
Attribute VB_Name = "arSalesSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pFrom As Date, pTO As Date, strCashSales As String, strCashSalesQty As String, _
    strDiscount As String, strTax As String)
    
    fCashSales.Text = strCashSales
    fQty.Text = strCashSalesQty
    fDiscount.Text = strDiscount
    fTax.Text = strTax
    Me.fReportTitle = "Sales summary between " & Format(pFrom, "dd/mm/yyyy HH:NN AMPM") & " and " & Format(pTO, "dd/mm/yyyy HH:NN AMPM")
    fEnd = "Printed : " & Format(Now, "dd/mm/yyyy HH:NN")
End Sub
