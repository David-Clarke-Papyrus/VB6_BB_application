VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arNielsenSales 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14925
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26326
   _ExtentY        =   12912
   SectionData     =   "arNielsenSales.dsx":0000
End
Attribute VB_Name = "arNielsenSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dblGroupTotal As Double
Public Sub Component(pLabel As String, pDatasource As ADODB.Recordset, pDate As Date)
    Set DC1.Recordset = pDatasource
    Me.lblHeading.Caption = pLabel
    dblGroupTotal = 0
End Sub

Private Sub Detail_Format()
    fValue.DataValue = fQty.DataValue * fPrice.DataValue
    dblGroupTotal = dblGroupTotal + fValue.DataValue
End Sub

Private Sub GroupFooter1_Format()
fExtTotal.DataValue = dblGroupTotal
End Sub
