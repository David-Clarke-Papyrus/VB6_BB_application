VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCashupExt_ForReportsApp 
   Caption         =   "ActiveReport1"
   ClientHeight    =   14175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15300
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26988
   _ExtentY        =   25003
   SectionData     =   "arCASHUPExt2_forReportsApp.dsx":0000
End
Attribute VB_Name = "arCashupExt_ForReportsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim strWithdrawals As String
Dim strCredits As String
Dim strNett As String
Dim oSQL As z_SQL


Public Sub Component(pRs As ADODB.Recordset, pCredits As String, pWithdrawals As String, pNett As String, Optional ZID As String)
    Set rs = pRs
    strWithdrawals = pWithdrawals
    strCredits = pCredits
    strNett = pNett
    
' '   Set arPC = New arCashup_PettyCash
' '   arPC.Component rs
''
'    Set oSQL = New z_SQL
'    oSQL.CalculateJournalSummary , , ZID
'    Dim rsS As ADODB.Recordset
'    Set rsS = New ADODB.Recordset
'    rsS.Open "SELECT * FROM tTJ", oPC.COShort
'    Set srDC1.Recordset = rsS
End Sub

'Private Sub ActiveReport_ReportStart()
'    Set sr1.Object = New arCashup_PettyCash
'    Set sr1.Object.DC1.Recordset = rs
'    sr1.Object.Sections("REPORTFOOTER").Controls.Item("fPCWithdrawals").Text = strWithdrawals
'    sr1.Object.Sections("REPORTFOOTER").Controls.Item("fPCCredits") = strCredits
'    sr1.Object.Sections("REPORTFOOTER").Controls.Item("fPCNett") = strNett
'End Sub



