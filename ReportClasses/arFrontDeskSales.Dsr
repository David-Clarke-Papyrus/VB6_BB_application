VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arFrontDeskSales 
   Caption         =   "Sales Report"
   ClientHeight    =   8235
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15075
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26591
   _ExtentY        =   14526
   SectionData     =   "arFrontDeskSales.dsx":0000
End
Attribute VB_Name = "arFrontDeskSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim oReport As z_reports
Dim strTPName As String
Dim strCode As String
'Dim strDate1 As String
'Dim strDate2 As String

''Dim lngTotal As Long
''Dim lngGrandTotal As Long
''Dim curTotal1 As Currency
''Dim curTotal2 As Currency
''Dim curGrandTotal1 As Currency
''Dim curGrandTotal2 As Currency

Public Sub Component(pRs As ADODB.Recordset, pCSCodeFrom As String, pCSCodeTo As String)
    Set rs = pRs
    
    lblRptHeader.Caption = "Front desk sales (by batch code) generated between " & pCSCodeFrom _
                                        & " and " & pCSCodeTo
    Set DC1.Recordset = pRs
    Me.lblPageHeader = "Printed on " & Format(Now(), "dd/mm/yyyy HH:NN")
    Me.lblFooterDate = ""
End Sub
Public Sub Component2(pRs As ADODB.Recordset, pDateFrom As String, pDateTo As String)
    Set rs = pRs
    
    lblRptHeader.Caption = "Front desk sales between " & pDateFrom _
                                        & " and " & pDateTo
    Set DC1.Recordset = pRs
    Me.lblPageHeader = "Printed on " & Format(Now(), "dd/mm/yyyy HH:NN")
    Me.lblFooterDate = ""
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 500
    Me.top = 200
    Me.Height = 7000
    Me.Width = 10000
    
''    ghTP.GroupValue = rs!TP_Name
''    txtghSupplier.Text = rs!TP_Name
End Sub

'''Private Sub Detail_AfterPrint()
'''    If rs.EOF Then Exit Sub
'''
'''    ghTP.GroupValue = rs!TP_Name
'''End Sub
'''
'''Private Sub Detail_Format()
'''On Error GoTo ERR_Handler
'''Dim cTmp1 As Currency
'''Dim cTmp2 As Currency
'''
'''    If rs.EOF Then GoTo EXIT_Handler
'''
'''    If HasNonEmptyString(rs!P_MainAuthor) Then
'''        txtDetails.Text = FNS(rs!P_Code) & " " & FNS(rs!P_Title) & vbCrLf & "Author:  " & FNS(rs!P_MainAuthor)
'''    Else
'''        txtDetails.Text = FNS(rs!P_Code) & " " & FNS(rs!P_Title)
'''    End If
'''
'''    If FNN(rs!TR_Type) = 0 Then
'''        txtTran.Text = ""
'''    Else
'''        txtTran.Text = FNN(rs!TR_Type)
'''    End If
'''    txtQty.Text = FNN(rs!CSL_Qty)
'''    cTmp1 = FNCURR(rs!CSL_Price) * FNN(rs!CSL_Qty) / 100
'''    txtPrice.Text = Format(cTmp1, "Standard")
'''    cTmp2 = (FNN(rs!CSL_Qty) * FNCURR(rs!CSL_Price)) * (1 - FNDBL(rs!CSL_Discount)) / 100
'''    txtNettPrice.Text = Format(cTmp2, "Standard")
'''
'''    lngTotal = lngTotal + FNN(rs!CSL_Qty)
'''    curTotal1 = curTotal1 + cTmp1
'''    curTotal2 = curTotal2 + cTmp2
'''
'''    Detail.PrintSection
'''    rs.MoveNext
'''
'''EXIT_Handler:
'''    Exit Sub
'''ERR_Handler:
'''    MsgBox Error
'''    GoTo EXIT_Handler
'''    Resume
'''End Sub
'''
'''Private Sub ghTP_Format()
'''    If rs.EOF Then Exit Sub
'''
'''    strTPName = rs!TP_Name
'''    txtghSupplier.Text = strTPName
'''End Sub
'''
'''Private Sub gfTP_AfterPrint()
'''    lngGrandTotal = lngGrandTotal + lngTotal
'''    curGrandTotal1 = curGrandTotal1 + curTotal1
'''    curGrandTotal2 = curGrandTotal2 + curTotal2
'''
'''    lngTotal = 0
'''    curTotal1 = 0
'''    curTotal2 = 0
'''End Sub
'''
'''Private Sub gfTP_Format()
'''    txtSubTotQty.Text = Format(lngTotal, "# ##0")
'''    txtSubTotPrice.Text = Format(curTotal1, "Currency")
'''    txtSubTotNettPrice.Text = Format(curTotal2, "Currency")
'''    txtSubTotal.Text = "Total for " & LCase(Trim$(strTPName))
'''End Sub
'''
'''Private Sub PageFooter_Format()
'''    lblFooterDate.Caption = Format(Date, "dddd, dd mmm yyyy")
'''End Sub
'''
'''Private Sub ReportFooter_Format()
'''    txtGrandTotQty.Text = Format(lngGrandTotal, "# ##0")
'''    txtGrandTotPrice.Text = Format(curGrandTotal1, "Currency")
'''    txtGrandTotalNettPrice.Text = Format(curGrandTotal2, "Currency")
'''End Sub
