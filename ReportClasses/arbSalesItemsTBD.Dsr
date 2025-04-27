VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesItems 
   Caption         =   "Sales Report"
   ClientHeight    =   8235
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15075
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26591
   _ExtentY        =   14526
   SectionData     =   "arbSalesItemsTBD.dsx":0000
End
Attribute VB_Name = "arSalesItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oReport As z_reports
Dim strTPName As String
Dim strCode As String
Dim strDate1 As String
Dim strDate2 As String

Dim lngTotal As Long
Dim lngGrandTotal As Long
Dim curTotal1 As Currency
Dim curTotal2 As Currency
Dim curGrandTotal1 As Currency
Dim curGrandTotal2 As Currency

Public Sub Component(pRS As ADODB.Recordset, pDate1 As Date, pDate2 As Date)
    Set rs = pRS
    
    lblRptHeader.Caption = "Products sold between " & Format(pDate1, "dd/mm/yyyy") _
                                        & " and " & Format(pDate2, "dd/mm/yyyy")
    Me.lblFooter.Caption = "Products sold"
    
    strDate1 = ReverseDate(pDate1)
    strDate2 = ReverseDate(DateAdd("d", 1, pDate2))
    Set DC1.Recordset = pRS
End Sub

Private Sub ActiveReport_Initialize()
    Set DC1.Connection = oPC.CO
    fSP.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
    fValu.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
    fNetValu.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
    fExVAT.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
 '   fSPT1.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
    fValuT1.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
    fNetValuT1.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
    fExVATT1.OutputFormat = oPC.Configuration.DefaultCurrency.FormatString
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 500
    Me.Top = 200
    Me.Height = 7000
    Me.Width = 10000
    
End Sub

Private Sub Detail_Format()
'On Error GoTo ERR_Handler
'Dim cTmp1 As Currency
'Dim cTmp2 As Currency
'
'    If rs.EOF Then GoTo EXIT_Handler
'
'    txtDetails = FNS(rs!Code) & " " & FNS(rs!Title)
'    txtTran = FNS(rs!SaleType)
'    txtQty.Text = FNN(rs!Qty)
'    fPrice = Format(rs!SP / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    fValu = Format(rs!Valu / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    fNetValu = Format(rs!NetValu / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    fExVat = Format(rs!ExVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'
'    lngTotal = lngTotal + FNN(rs!Qty)
'    curTotal1 = curTotal1 + cTmp1
'    curTotal2 = curTotal2 + cTmp2
'
'    Detail.PrintSection
'    rs.MoveNext
    fSP.DataValue = fSP.DataValue / oPC.Configuration.DefaultCurrency.Divisor
    fValu.DataValue = fValu.DataValue / oPC.Configuration.DefaultCurrency.Divisor
    fNetValu.DataValue = CLng(fNetValu.DataValue) / oPC.Configuration.DefaultCurrency.Divisor
    fExVAT.DataValue = CLng(fExVAT.DataValue) / oPC.Configuration.DefaultCurrency.Divisor
EXIT_Handler:
    Exit Sub
ERR_Handler:
    MsgBox Error
    GoTo EXIT_Handler
    Resume
End Sub

'Private Sub ghTP_Format()
'    If rs.EOF Then Exit Sub
'
'    strTPName = rs!TP_Name
'    txtghSupplier.Text = strTPName
'End Sub
'
'Private Sub gfTP_AfterPrint()
'    lngGrandTotal = lngGrandTotal + lngTotal
'    curGrandTotal1 = curGrandTotal1 + curTotal1
'    curGrandTotal2 = curGrandTotal2 + curTotal2
'
'    lngTotal = 0
'    curTotal1 = 0
'    curTotal2 = 0
'End Sub
'
'Private Sub gfTP_Format()
'    txtSubTotQty.Text = Format(lngTotal, "# ##0")
'    txtSubTotPrice.Text = Format(curTotal1, "Currency")
'    txtSubTotNettPrice.Text = Format(curTotal2, "Currency")
'    txtSubTotExVAT.Text = Format(curTotal2, "Currency")
'    txtSubTotal.Text = "Total for " & LCase(Trim$(strTPName))
'End Sub

Private Sub PageFooter_Format()
    lblFooterDate.Caption = Format(Date, "dddd, dd mmm yyyy")
End Sub

Private Sub ReportFooter_Format()
  '  Me.fSPT1.DataValue = fSPT1.DataValue / oPC.Configuration.DefaultCurrency.Divisor
    fValuT1.DataValue = fValuT1.DataValue / oPC.Configuration.DefaultCurrency.Divisor
    fNetValuT1.DataValue = CLng(fNetValuT1.DataValue) / oPC.Configuration.DefaultCurrency.Divisor
    fExVATT1.DataValue = CLng(fExVATT1.DataValue) / oPC.Configuration.DefaultCurrency.Divisor
End Sub
