VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arAppro 
   Caption         =   "Value of stock out on appro"
   ClientHeight    =   9780
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   _ExtentX        =   26882
   _ExtentY        =   17251
   SectionData     =   "arAppro2.dsx":0000
End
Attribute VB_Name = "arAppro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As ADODB.Recordset

Dim cTotalCost As Currency
Dim cSubTotalCost As Currency
Dim cGrandTotalCost As Currency

Dim cGRTotal As Currency
Dim cGRSubTotal As Currency
Dim cGRGrandTotal As Currency

Dim cNettTotal As Currency
Dim cNettSubTotal As Currency
Dim cNettGrandTotal As Currency

Dim lngSupplierQtyTotal As Long
Dim lngTRCodeQtyTotal As Long
Dim lngGrandQty As Long

Dim oReport As z_reports
Dim strTPName As String
Dim strTRCode As String
Dim blnSuppChange As Boolean
Dim blnAll As Boolean

Public Sub Component(pRs As ADODB.Recordset, pTitle As String, pFooter As String, Optional pUsesLPD As Boolean)
    Set DC1.Recordset = pRs
    Me.lblHeader.Caption = pTitle
    Me.lblFooter.Caption = pFooter
    Me.lblFooterDate = Format(Now(), "DD-MM-YYYY Hh:Nn")
    If pUsesLPD Then lblCost.Caption = "Value at last del. cost (Ex VAT)"
End Sub

