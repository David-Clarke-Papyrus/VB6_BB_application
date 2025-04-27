VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesByCustomer1 
   Caption         =   "Sales by Customer"
   ClientHeight    =   7770
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   19125
   _ExtentX        =   33734
   _ExtentY        =   13705
   SectionData     =   "arSalesByCustomer1.dsx":0000
End
Attribute VB_Name = "arSalesByCustomer1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Public Sub Component(pRS As ADODB.Recordset, pFrom As Date, pTo As Date, pstrCostSetting As String)
    Set rs = pRS
    Set DataControl1.Recordset = rs
    lblRptHeader.Caption = "Sales summary by customer from " & Format(pFrom, "dd/mm/yyyy") _
                                        & " to " & Format(pTo, "dd/mm/yyyy")
    Me.lblCostNote.Caption = IIf(pstrCostSetting = "LPD", "Cost uses last delivered price (ex VAT)", "Cost is weighted average cost (ex VAT)")
End Sub



Private Sub ReportFooter_BeforePrint()
    If fTotalSP.DataValue <> 0 Then
        ftotalGMPer.DataValue = CCur(fTotalGM.DataValue) / CCur(fTotalSP.DataValue)
    Else
        ftotalGMPer.Text = ""
    End If

End Sub

