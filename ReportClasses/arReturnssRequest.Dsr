VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arReturnsRequest 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arReturnssRequest.dsx":0000
End
Attribute VB_Name = "arReturnsRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long

Sub component(pRs As ADODB.Recordset, pSupplierName As String, pDocCOde As String)
    Set rs = pRs
    lngRC = rs.RecordCount
    tDatePrinted = "Printed: " & Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 13000
    Me.Height = 6000
    i = 1
    fFROM = oPC.Configuration.DefaultStore.BillAddress
    lblPH.Caption = "Return to : " & pSupplierName & "   Document number : " & pDocCOde
End Sub



Private Sub Detail_Format()
    If rs.eof Then Exit Sub
    tTitle = rs.Fields("Title")
    tCode = rs.Fields("Code")
    tPubcode = rs.Fields("PubCode")
    tQty = rs.Fields("Requested")
    tRef = rs.Fields("SupplierInvoiceRef")
    rs.MoveNext
    Detail.PrintSection
End Sub



