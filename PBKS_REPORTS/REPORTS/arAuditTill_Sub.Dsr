VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arAuditTill_Sub 
   Caption         =   "ActiveReport1"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19920
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35137
   _ExtentY        =   14923
   SectionData     =   "arAuditTill_Sub.dsx":0000
End
Attribute VB_Name = "arAuditTill_Sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Public Sub component(pFrom As Date, pTo As Date, Optional ZID As String)
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If ZID > "" Then
    rs.Open "SELECT SUM(IsPriceAlteration) AS PriceAlterations, SUM(IsDiscounted) AS Discounts, SUM(IsSoldWhenNegative) AS NegativeStock, SUM(IsReturn) " _
            & " AS [Returns], SUM(IsVoider) AS Voider, SUM(IsVoided) AS Voided " _
            & " FROM dbo.ahv_TillPointAudit WHERE Z_ID = '" & ZID & "'", oPC.CO, adOpenForwardOnly, adLockOptimistic
    Else
    rs.Open "SELECT SUM(IsPriceAlteration) AS PriceAlterations, SUM(IsDiscounted) AS Discounts, SUM(IsSoldWhenNegative) AS NegativeStock, SUM(IsReturn) " _
            & " AS [Returns], SUM(IsVoider) AS Voider, SUM(IsVoided) AS Voided " _
            & " FROM dbo.ahv_TillPointAudit WHERE ZSTART > = {d '" & ReverseDate(pFrom) & "'} AND ZSTART <=  {d '" & ReverseDate(pTo) & "'}", oPC.CO, adOpenForwardOnly, adLockOptimistic
    End If
End Sub

Private Sub ActiveReport_DataInitialize()
    Set Me.srDC1.Recordset = rs
End Sub

Private Sub ActiveReport_ReportEnd()
    rs.Close
    Set rs = Nothing
End Sub

