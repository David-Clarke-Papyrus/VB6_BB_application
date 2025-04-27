VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arXReport 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24844
   _ExtentY        =   14393
   SectionData     =   "arXReport.dsx":0000
End
Attribute VB_Name = "arXReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsX As ADODB.Recordset
Sub Component(prsX As ADODB.Recordset, pHeading As String)
    Set rs = prsX
    Me.lblReport.Caption = pHeading
End Sub

Private Sub ActiveReport_ReportStart()
    DC1.Recordset = rsX
End Sub

Private Sub Detail_BeforePrint()

End Sub

    Me.Sections(1).Controls("fAmt_GTot") = Format(CCur(srInv.object.fAMT_TOT) - CCur(srCN.object.fAMT_TOT), "###,##0.00")
    Me.Sections(1).Controls("fVAT_GTOT") = Format(CCur(srInv.object.fVAT_TOT) - CCur(srCN.object.fVAT_TOT), "###,##0.00")
    Me.Sections(1).Controls("fEXCLAMT_GTOT") = Format(CCur(srInv.object.fEXCLAMT_TOT) - CCur(srCN.object.fEXCLAMT_TOT), "###,##0.00")

Private Sub ActiveReport_ReportEnd()
    On Error Resume Next
    rsX.Close
    Set rsX = Nothing
End Sub


