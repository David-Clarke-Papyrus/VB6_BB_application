VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSupplierClaim 
   Caption         =   "Supplier claim"
   ClientHeight    =   9465
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   19080
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   33655
   _ExtentY        =   16695
   SectionData     =   "arSupplierClaim.dsx":0000
End
Attribute VB_Name = "arSupplierClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long
Dim dblSupplierTotal As Double
Dim dblGrandTotal As Double

Sub Component(pRS As ADODB.Recordset)
    Set rs = pRS
    DC1.Recordset = rs
    rs.Sort = "LASTSUPPLIERNAME"
'    fSupplierName.DataField = "LASTSUPPLIERNAME"
'    SupplierHead.DataField = "LASTSUPPLIERNAME"
'    If pFilter Then
'        rs.Filter = "QTYFIRM > 0 or QTYSS > 0"
'    End If
    lngRC = rs.RecordCount
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn AM/PM")
'    If Me.WindowState <> 2 Then
'        Me.Width = 12000
'        Me.Height = 6000
'        Me.top = 2000
'        Me.Left = 1000
'    End If
    i = 1
End Sub



Private Sub ActiveReport_FetchData(eof As Boolean)
'If DC1.Recordset.eof Then Exit Sub
'    Fields("Extended").Value = (FNN(DC1.Recordset.Fields("QtyFirm")) + FNN(DC1.Recordset.Fields("QtySS"))) * (FNN(DC1.Recordset.Fields("PRICE")) / oPC.Configuration.DefaultCurrency.Divisor)

End Sub

Private Sub ActiveReport_Initialize()
Dim dteTMP As Date
    tSInvoice.DataField = "SupplierInvoiceCode"
    tSDate.DataField = "PRCODE"
    tGRNDoc.DataField = "ONHAND"
    tDocDate.DataField = "QTYCO"
    tISBN.DataField = "QTYPO"
    tPrice.DataField = "QTYAPP"
    tShrtDam.DataField = "LASTSIXWEEKS"
    tCorrDisc.DataField = "LASTSIXMONTHS"
    tCPrice.DataField = "TOTALSOLD"
    tExplanation.DataField = "PUBLISHER"
    tClaimValue.DataField = "Extended"
    Me.fTotal.DataField = "EXTENDED"
End Sub

