VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCashupEx 
   Caption         =   "ActiveReport1"
   ClientHeight    =   13620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14145
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   24950
   _ExtentY        =   24024
   SectionData     =   "arCASHUPEx.dsx":0000
End
Attribute VB_Name = "arCashupEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim strWithdrawals As String
Dim strCredits As String
Dim strNett As String


Public Sub Component(pRs As ADODB.Recordset, pCredits As String, pWithdrawals As String, pNett As String)
    Set rs = pRs
    strWithdrawals = pWithdrawals
    strCredits = pCredits
    strNett = pNett
End Sub

Private Sub ActiveReport_ReportStart()
  '  Set SR1.object = New arCashup_PettyCash
 '  SR1.object.Component rs
End Sub

Private Sub Detail_Format()
   ' Set Me.SR1.object.DC1.Recordset = rs
'    SR1.object.DC1.Refresh
'    If rs.Fields.Count > 0 Then
'        SR1.object.Sections("DETAIL").Controls(1) = rs.Fields(1)
'    End If
  '  SR1.object.Sections("REPORTFOOTER").Controls("fPCWithdrawals") = strWithdrawals
  '  SR1.object.Sections("REPORTFOOTER").Controls("fPCCredits") = strCredits
  '  SR1.object.Sections("REPORTFOOTER").Controls("fPCNett") = strNett
End Sub
