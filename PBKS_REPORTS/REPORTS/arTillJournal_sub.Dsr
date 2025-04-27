VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arTillJournal_sub 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10425
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   18389
   SectionData     =   "arTillJournal_sub.dsx":0000
End
Attribute VB_Name = "arTillJournal_sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oSQL As z_SQL
Public Sub Component(pFrom As Date, pTo As Date)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set oSQL = New z_SQL
    oSQL.CalculateJournalSummary pFrom, pTo
    rs.Open "SELECT * FROM tTJ", oPC.CO, adOpenForwardOnly, adLockOptimistic
End Sub

Private Sub ActiveReport_DataInitialize()
    Set Me.srDC1.Recordset = rs
End Sub

Private Sub ActiveReport_ReportEnd()
    rs.Close
    Set rs = Nothing
End Sub



