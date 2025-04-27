VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesByCustomerBudget 
   Caption         =   "Sales by Customer"
   ClientHeight    =   7770
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   18960
   MDIChild        =   -1  'True
   _ExtentX        =   33443
   _ExtentY        =   13705
   SectionData     =   "arSalesByCustomerBudget.dsx":0000
End
Attribute VB_Name = "arSalesByCustomerBudget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Public Sub Component(pRs As ADODB.Recordset, pFrom As Date, pTo As Date)
    Set rs = pRs
    Set DataControl1.Recordset = rs
    lblRptHeader.Caption = "Sales summary by customer from " & Format(pFrom, "dd/mm/yyyy") _
                                        & " to " & Format(pTo, "dd/mm/yyyy")
    
    Me.Visible = False
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Height = 7000
    Me.Width = 13000
    Me.Visible = True
    
End Sub

