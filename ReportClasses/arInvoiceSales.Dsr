VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arInvoiceSales 
   Caption         =   "Sales Report"
   ClientHeight    =   11115
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19606
   SectionData     =   "arInvoiceSales.dsx":0000
End
Attribute VB_Name = "arInvoiceSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oReport As z_reports
Dim strTPName As String
Dim strCode As String

Public Sub Component(pRs As ADODB.Recordset, pHeading As String)
    Set rs = pRs
    
    lblRptHeader.Caption = pHeading
    Me.lblFooter.Caption = "Invoice sales"
    Set DC1.Recordset = pRs
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 500
    Me.top = 200
    Me.Height = 7000
    Me.Width = 10000
    
End Sub

