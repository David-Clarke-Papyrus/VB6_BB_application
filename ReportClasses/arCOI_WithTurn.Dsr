VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCOI_WithTurn 
   Caption         =   "Cost of inventory"
   ClientHeight    =   10800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   19560
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   34502
   _ExtentY        =   19050
   SectionData     =   "arCOI_WithTurn.dsx":0000
End
Attribute VB_Name = "arCOI_WithTurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mUsesLPD As Boolean

Public Sub Component(pRS As ADODB.Recordset, pMsg As String, pUsesLPD As Boolean, ByWhat As String)
    mUsesLPD = pUsesLPD
    Set DC1.Recordset = pRS
    fReportTitle = "Cost of Inventory          printed: " & Format(Now(), "dd/mm/yyyy")
    lblNote.Caption = pMsg
    
    lblNote2.Caption = "* " & IIf(pUsesLPD = True, "  calculated using last delivered cost", "")
    fEND = "Printed : " & Format(Now, "dd/mm/yyyy HH:NN")
    If ByWhat = "PT" Then
        lblBD.Caption = "Product type"
    ElseIf ByWhat = "C" Then
        lblBD.Caption = "Category"
    ElseIf ByWhat = "P" Then
        lblBD.Caption = "Publisher"
    End If
End Sub





Private Sub ActiveReport_ReportStart()
    If mUsesLPD Then
        f12.DataField = "OHLPD"
        fGT12.DataField = "OHLPD"
        f23.DataField = "APP_OS_GM_LDP"
        f22.DataField = "APP_OS_LPD"
        fGT13.DataField = "OHGM_LDP"
        fGT22.DataField = "APP_OS_LPD"
        fGT23.DataField = "APP_OS_GM_LDP"
        f2GT.DataField = "GTAPPGM_LDP"
        f1GT.DataField = "GTOHGM_LDP"
        f13.DataField = "OHGM_LDP"
    Else
    End If
    f11.OutputFormat = "#,##0"
    f12.OutputFormat = "#,##0"
    f13.OutputFormat = "#,##0"
    f21.OutputFormat = "#,##0"
    f22.OutputFormat = "#,##0"
    f23.OutputFormat = "#,##0"
    f31.OutputFormat = "#,##0"
    f32.OutputFormat = "#,##0"
    f33.OutputFormat = "#,##0"
    f14.OutputFormat = "#.00%"
    f24.OutputFormat = "#.00%"
    f34.OutputFormat = "#.00%"
    fGT11.OutputFormat = "#,##0"
    fGT12.OutputFormat = "#,##0"
    fGT13.OutputFormat = "#,##0"
    fGT21.OutputFormat = "#,##0"
    fGT22.OutputFormat = "#,##0"
    fGT23.OutputFormat = "#,##0"
    fGT31.OutputFormat = "#,##0"
    fGT32.OutputFormat = "#,##0"
    fGT33.OutputFormat = "#,##0"
End Sub

Private Sub GroupFooter1_Format()
End Sub

Private Sub Detail_Format()
    f14.DataValue = 0
    f24.DataValue = 0
    f34.DataValue = 0
    If f1GT.DataValue > 0 Then f14.DataValue = FNDBL((f13.DataValue / f1GT.DataValue)) ' * 100
    If f2GT.DataValue > 0 Then f24.DataValue = FNDBL((f23.DataValue / f2GT.DataValue)) ' * 100
    If f3GT.DataValue > 0 Then f34.DataValue = FNDBL((f33.DataValue / f3GT.DataValue)) ' * 100
End Sub
