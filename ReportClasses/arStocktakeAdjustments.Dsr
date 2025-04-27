VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arStocktakeAdjustments 
   Caption         =   "arMissing"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17370
   StartUpPosition =   2  'CenterScreen
   _ExtentX        =   30639
   _ExtentY        =   11060
   SectionData     =   "arStocktakeAdjustments.dsx":0000
End
Attribute VB_Name = "arStocktakeAdjustments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub Component(pRs As ADODB.Recordset)
    Set rs = pRs
    lblReportHeader.Caption = "Stock discrepancies for stock-take with cut-off date: " & Format(rs.Fields("STKTKE_CUTOFFDATE"), "dd/mm/yyyy HH:NN")
    Me.Caption = lblReportHeader.Caption
    Width = 10000
    Height = 5000
End Sub
Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    txtCode = FNS(rs.Fields("P_Code"))
    txtTitle = FNS(rs.Fields("P_Title"))
    txtOnHand = FNN(rs.Fields("P_QtyOnHand"))
    txtPrice = Format(FNN(rs.Fields("P_RRP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    txtQty = FNN(rs.Fields("STKTKEL_Difference"))
    txtRRPVal = Format(FNN(rs.Fields("RRPVAL")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    txtCostVal = Format(FNN(rs.Fields("COSTVAL")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    txtCost = Format(FNN(rs.Fields("P_COST")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    rs.MoveNext
    Detail.PrintSection
End Sub

