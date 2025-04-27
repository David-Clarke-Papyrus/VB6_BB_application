VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSupplierClaim 
   Caption         =   "Claim on Vendor"
   ClientHeight    =   9150
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   20295
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   35798
   _ExtentY        =   16140
   SectionData     =   "arSupplierClaimb.dsx":0000
End
Attribute VB_Name = "arSupplierClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsClaim As ADODB.Recordset

Dim lngCount As Long

Public Sub component(rs As ADODB.Recordset, pSupplierName As String, pClaimRef As String, TotalClaimValue As String)
   ' rs.MoveFirst
    Set DC1.Recordset = rs
    tStore = oPC.Configuration.DefaultStore.BillAddress
    tSupplier = pSupplierName
    tClaimRef = pClaimRef
    tDeliveredTo = oPC.Configuration.DefaultStore.DelAddress
    Me.tTotalValue = FNS(rs.Fields("TotalClaimValueF"))
End Sub


Private Sub PageHeader_Format()
    Shape2.Height = tDeliveredTo.Height + 1250
End Sub
