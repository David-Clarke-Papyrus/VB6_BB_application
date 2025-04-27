VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arValidation_D 
   Caption         =   "arMissing"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14985
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26432
   _ExtentY        =   11060
   SectionData     =   "arValidation_D.dsx":0000
End
Attribute VB_Name = "arValidation_D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset

Public Sub Component(pRS As ADODB.Recordset, pTitle As String)
    Set rs = pRS
    Me.lblReportHeader = pTitle
    tDte = "Printed : " & Format(Now, "dd/mm/yyyy HH:NN")
End Sub
Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    Me.txtCode = FNS(rs.Fields("P_Code"))
    Me.txtTitle = FNS(rs.Fields("P_Title"))
    Me.txtOnHand = FNN(rs.Fields("P_QtyOnHand"))
    Me.txtPrice = Format(FNN(rs.Fields("P_SP")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Me.txtQty = FNN(rs.Fields("STKTKEL_Difference"))
    Me.fCost = Format(FNN(rs.Fields("P_COST")) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    Me.txtID = rs.Fields("STKTKEL_ID")
    rs.MoveNext
    Detail.PrintSection
End Sub

