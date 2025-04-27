VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCOI_STOCKTAKE_ADJ 
   Caption         =   "Cost of inventory"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16200
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   28575
   _ExtentY        =   17304
   SectionData     =   "arCOI_STOCKTAKE_ADJ.dsx":0000
End
Attribute VB_Name = "arCOI_STOCKTAKE_ADJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mUsesLPD As Boolean

Public Sub Component(pRs As ADODB.Recordset, pDate As String, pMsg As String, pUsesLPD As Boolean)
    mUsesLPD = pUsesLPD
    Set DC1.Recordset = pRs
    fReportTitle = "Cost of Inventory  -  printed: " & Format(Now(), "dd/mm/yyyy") & " for Stock-take at " & pDate
    lblNote.Caption = pMsg & IIf(pUsesLPD = True, "  calculated using last delivered cost", "")
    fEND = "Printed : " & Format(Now, "dd/mm/yyyy HH:NN")
End Sub





Private Sub ActiveReport_ReportStart()
    If mUsesLPD Then
        f12.DataField = "OHLPD"
        fGT12.DataField = "OHLPD"
    Else
    End If
    f11.OutputFormat = "#,##0"
    f12.OutputFormat = "#,##0"
    f13.OutputFormat = "#,##0"
    f14.OutputFormat = "#.00%"
    
    fGT11.OutputFormat = "#,##0"
    fGT12.OutputFormat = "#,##0"
    fGT13.OutputFormat = "#,##0"
End Sub

Private Sub GroupFooter1_Format()
End Sub

Private Sub Detail_Format()
    If f1GT.DataValue <> 0 Then f14.DataValue = FNDBL((f13.DataValue / f1GT.DataValue))
End Sub
