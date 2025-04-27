VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arPOReminder 
   Caption         =   "Purchase order reminder previewer"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arPOReminder.dsx":0000
End
Attribute VB_Name = "arPOReminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub component(pRs As ADODB.Recordset, pPagePerSupplier As Boolean)
    Set rs = pRs
    DC1.Recordset = rs
    SupplierHeader.DataField = "TP_Name"
    If pPagePerSupplier Then
        SUpplierFooter.NewPage = ddNPBefore
    End If
    lngRC = rs.RecordCount
    tDatePrinted = Format(Now(), "dd-mm-yyyy hh:nn")
    lblTitle.Caption = "Reminder from " & oPC.Configuration.DefaultCompany.CompanyName
    Me.Width = 8000
    Me.Height = 8000
    Me.PrintWidth = 10200
   i = 1
End Sub
Private Sub ActiveReport_Initialize()
'    tTitle.DataField = "P_Title"
'    fCode.DataField = "P_Code"
'    tDate.DataField = "TR_Date"
'    tRecd.DataField = "POL_QtyReceivedSoFar"
    tSupplier.DataField = "TP_Name"
End Sub



Private Sub Detail_Format()
    If DC1.Recordset.Fields("POL_REF") > "" Then
        tCode = FNS(DC1.Recordset.Fields("POL_REF"))
    Else
        tCode = DC1.Recordset.Fields("TR_Code")
    End If
'    tQty = DC1.Recordset.Fields("POL_QTYFirm") & "/" & DC1.Recordset.Fields("POL_QtySS")
'    tOS = FNN(DC1.Recordset.Fields("POL_QTYFirm")) + FNN(DC1.Recordset.Fields("POL_QtySS")) - FNN(rs.Fields("POL_QtyReceivedSoFar"))
'    tAction = ""
End Sub


