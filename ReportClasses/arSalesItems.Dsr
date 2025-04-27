VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arSalesItems 
   Caption         =   "Sales Report"
   ClientHeight    =   6600
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12480
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22013
   _ExtentY        =   11642
   SectionData     =   "arSalesItems.dsx":0000
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

Public Sub Component(pRs As ADODB.Recordset, pDate1 As Date, pDate2 As Date)
    Set rs = pRs
    
    lblRptHeader.Caption = "Products sold between " & Format(pDate1, "dd/mm/yyyy") _
                                        & " and " & Format(pDate2, "dd/mm/yyyy")
    Me.lblFooter.Caption = "Products sold"
    
    strDate1 = ReverseDate(pDate1)
    strDate2 = ReverseDate(DateAdd("d", 1, pDate2))
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 500
    Me.top = 200
    Me.Height = 7000
    Me.Width = 10000
    
End Sub

Private Sub Detail_Format()
On Error GoTo Err_Handler
Dim cTmp1 As Currency
Dim cTmp2 As Currency
    
    If rs.EOF Then GoTo EXIT_Handler
            
    txtDetails = FNS(rs!code) & " " & FNS(rs!Title)
    txtTran = FNS(rs!SaleType)
    txtQty.text = FNN(rs!Qty)
    fPrice = Format(rs!SP / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    fValu = Format(rs!Valu / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    fNetValu = Format(rs!NetValu / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    fExVat = Format(rs!ExVAT / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
    
    lngTotal = lngTotal + FNN(rs!Qty)
    curTotal1 = curTotal1 + cTmp1
    curTotal2 = curTotal2 + cTmp2
    
    Detail.PrintSection
    rs.MoveNext
    
EXIT_Handler:
    Exit Sub
Err_Handler:
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
    txtGrandTotalQty.text = Format(lngGrandTotal, "# ##0")
    txtGrandTotalPrice.text = Format(curGrandTotal1, "Currency")
    txtGrandTotalExVAT.text = Format(curGrandTotal1, "Currency")
    txtGrandTotalNetValu.text = Format(curGrandTotal2, "Currency")
    txtGrandTotalValu.text = Format(curGrandTotal2, "Currency")
End Sub
