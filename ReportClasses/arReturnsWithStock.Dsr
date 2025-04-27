VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arReturnWithStock 
   Caption         =   "ActiveReport1"
   ClientHeight    =   7935
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   1  'CenterOwner
   _ExtentX        =   26882
   _ExtentY        =   13996
   SectionData     =   "arReturnsWithStock.dsx":0000
End
Attribute VB_Name = "arReturnWithStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim i As Long
Dim lngRC As Long
Dim lngTotalQty As Long


Sub component(pRs As ADODB.Recordset, pSupplierName As String, pDocCOde As String, pApprovalRf As String, pTotalPayableF As String)
    
    Set rs = pRs
    lngRC = rs.RecordCount
    tDatePrinted = "Printed: " & Format(Now(), "dd-mm-yyyy hh:nn")
    Me.Width = 13000
    Me.Height = 6000
    i = 1
    fFROM = oPC.Configuration.DefaultStore.BillAddress
    lblPH.Caption = "Return to : " & pSupplierName & "   Document number : " & pDocCOde & "   Approval number : " & pApprovalRf
    Me.fTotalValue = pTotalPayableF
    lngTotalQty = 0
    Set DC1.Recordset = pRs
End Sub



Private Sub ActiveReport_Initialize()
    fValue.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
 '   fTotalValue.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
End Sub


'Private Sub Detail_AfterPrint()
'    If rs.EOF Then Exit Sub
'End Sub

Private Sub Detail_Format()
'    If rs.EOF Then Exit Sub
'    tTitle = rs.Fields("Title")
'    tCode = rs.Fields("Code")
'    tPubcode = rs.Fields("PubCode")
'    tQty = rs.Fields("Returned")
'    tRef = rs.Fields("SupplierInvoiceRef")
'    rs.MoveNext
'    Detail.PrintSection
'    fValue.DataValue = CLng(fValue.DataValue) / oPC.Configuration.DefaultCurrency.Divisor
End Sub



'Private Sub ReportFooter_Format()
'    Me.fTotalQty = lngTotalQty
'  '  Me.fTotalQty = rs.Fields("TotalQty")
'End Sub
Private Sub ReportFooter_Format()
'    fTotalValue.DataValue = CLng(fTotalValue.DataValue) / oPC.Configuration.DefaultCurrency.Divisor
End Sub
