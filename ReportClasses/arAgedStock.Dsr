VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arAgedStock 
   Caption         =   "Unordered deliveries"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15225
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26855
   _ExtentY        =   14393
   SectionData     =   "arAgedStock.dsx":0000
End
Attribute VB_Name = "arAgedStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Component(pRs As ADODB.Recordset, pdte1 As Date, pdte2 As Date, pUseLPD As Boolean)
    Set DC1.Recordset = pRs
    Me.Caption = "Aged stock"
    fReportTitle = "Aged stock received between " & Format(pdte1, "dd/mm/yyyy") & " and " & Format(pdte2, "dd/mm/yyyy")
    fEND = "Printed : " & Format(Now, "dd/mm/yyyy HH:NN")
    lblNote2.Caption = IIf(pUseLPD = True, "* calculated using last delivered cost (Ex VAT)", "* calculated using weighted average cost (Ex VAT)")
    lblCostDescription.Caption = IIf(pUseLPD = True, "* calculated using last delivered cost (Ex VAT)", "* calculated using weighted average cost (Ex VAT)")
End Sub
