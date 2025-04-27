VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arInvoices_CNs 
   Caption         =   "ActiveReport1"
   ClientHeight    =   12630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18825
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   33205
   _ExtentY        =   22278
   SectionData     =   "arInvoices_CNs.dsx":0000
End
Attribute VB_Name = "arInvoices_CNs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsINV As ADODB.Recordset
Dim rsCN As ADODB.Recordset
Sub Component(prsINV As ADODB.Recordset, prsCN As ADODB.Recordset, pHeading As String)
    Set rsINV = prsINV
    Set rsCN = prsCN
    Me.lblReport.Caption = pHeading
End Sub

Private Sub ActiveReport_ReportStart()
    Set srInv.Object = New arInv_Summ
    Set srCN.Object = New arCN_Summ
End Sub

Private Sub Detail_BeforePrint()
'    Me.Sections("Detail").Controls("fAmt_GTot") = Format(CCur(FNDBL(srInv.object.Sections("ReportFooter").Controls("fAMT_TOT"))) - CCur(FNDBL(srCN.object.Sections("ReportFooter").Controls("fAMT_TOT"))), "###,##0.00")
'    Me.Sections("Detail").Controls("fVAT_GTOT") = Format(CCur(FNDBL(srInv.object.Sections("ReportFooter").Controls("fVAT_TOT"))) - CCur(FNDBL(srCN.object.Sections("ReportFooter").Controls("fVAT_TOT"))), "###,##0.00")
'    Me.Sections("Detail").Controls("fEXCLAMT_GTOT") = Format(CCur(FNDBL(srInv.object.Sections("ReportFooter").Controls("fEXCLAMT_TOT"))) - CCur(FNDBL(srCN.object.Sections("ReportFooter").Controls("fEXCLAMT_TOT"))), "###,##0.00")
  '  Me.Sections("Detail").Controls("fVAT_GTOT") = Format(CCur(FNDBL(srInv.object.fVAT_TOT)) - CCur(FNDBL(srCN.object.fVAT_TOT)), "###,##0.00")
  '  Me.Sections("Detail").Controls("fEXCLAMT_GTOT") = Format(CCur(FNDBL(srInv.object.fEXCLAMT_TOT)) - CCur(FNDBL(srCN.object.fEXCLAMT_TOT)), "###,##0.00")

End Sub

Private Sub Detail_Format()
    Set srInv.Object.DC1.Recordset = rsINV
    Set srCN.Object.DC1.Recordset = rsCN
End Sub

Private Sub ActiveReport_ReportEnd()
    On Error Resume Next
    Unload srInv.Object
    Set srInv.Object = Nothing

    Unload srCN.Object
    Set srCN.Object = Nothing
End Sub


Private Sub ReportFooter_Format()
    Me.Sections("ReportFooter").Controls("fAmt_GTot") = Format(CCur(FNDBL(srInv.Object.Sections("ReportFooter").Controls("fAMT_TOT"))) - CCur(FNDBL(srCN.Object.Sections("ReportFooter").Controls("fAMT_TOT"))), "###,##0.00")
    Me.Sections("ReportFooter").Controls("fVAT_GTOT") = Format(CCur(FNDBL(srInv.Object.Sections("ReportFooter").Controls("fVAT_TOT"))) - CCur(FNDBL(srCN.Object.Sections("ReportFooter").Controls("fVAT_TOT"))), "###,##0.00")
    Me.Sections("ReportFooter").Controls("fEXCLAMT_GTOT") = Format(CCur(FNDBL(srInv.Object.Sections("ReportFooter").Controls("fEXCLAMT_TOT"))) - CCur(FNDBL(srCN.Object.Sections("ReportFooter").Controls("fEXCLAMT_TOT"))), "###,##0.00")
End Sub
