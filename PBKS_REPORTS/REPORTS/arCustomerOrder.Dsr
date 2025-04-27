VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arCustomerOrder 
   Caption         =   "Customer Orders - Summary"
   ClientHeight    =   9900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15255
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26908
   _ExtentY        =   17463
   SectionData     =   "arCustomerOrder.dsx":0000
End
Attribute VB_Name = "arCustomerOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim rsSub As ADODB.Recordset

Dim lngTPID As Long

Dim curQtyBooks As Currency
Dim curQtyTitles As Currency
Dim curVATExcl As Currency
Dim curVATIncl As Currency

Dim strType As String
Dim strQuery2 As String
Dim strDate1 As String
Dim strDate2 As String
Dim strWHERE As String
Dim strRptHeader As String

Public Sub Component(pRS As ADODB.Recordset, ByVal pRptHeader As String, ByVal pType As String, ByVal pDate1 As Date, _
                                    ByVal pDate2 As Date)
    Set rs = pRS
    strRptHeader = pRptHeader
    strType = pType
    strDate1 = ReverseDate(pDate1)
    If pDate2 > CDate(0) Then strDate2 = ReverseDate(DateAdd("d", 1, pDate2))
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.Top = 500
    Me.Height = 7000
    Me.Width = 10000
   
    txtReportHeader.Text = strRptHeader
    ghCustomer.GroupValue = rs!TP_Name
    PageHeader.Height = 0
End Sub

Private Sub ActiveReport_Terminate()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set SubRpt.object = Nothing
End Sub

Private Sub Detail_AfterPrint()
    If rs.EOF Then Exit Sub
    
    ghCustomer.GroupValue = rs!TP_Name
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    txtCustomer.Text = rs!TP_Name
    txtQtyBooks.Text = Format(rs!QtyBooks, "# ##0")
    txtQtyTitles.Text = Format(rs!QtyTitles, "# ##0")
    txtValInclVAT.Text = Format((rs!val / 100), "Currency")
    txtValExVAT.Text = Format((rs!val / 100) / (1 + (oPC.Configuration.VATRate / 100)), "Currency")
    txtDiscValInclVAT.Text = Format((rs!DiscountedVal / 100), "Currency")
    txtDiscValExVAT.Text = Format((rs!DiscountedVal / 100) / (1 + (oPC.Configuration.VATRate / 100)), "Currency")
    
    lngTPID = rs!TP_ID
    
    Me.Detail.PrintSection
    rs.MoveNext
End Sub

Private Sub gfCustomer_AfterPrint()
    If rs.EOF Then Exit Sub
    
    lngTPID = 0
End Sub

Private Sub gfCustomer_Format()
Dim rsSub As New ADODB.Recordset
Dim strSQL As String
    
    Set rsSub = Nothing
    Set SubRpt.object = New arCustomerOrderSub
    
    strWHERE = ""
    Select Case strType
    Case "Between"
        strWHERE = "((TR_Date) >= '" & strDate1 & "') AND ((TR_Date) < '" & strDate2 & "')"
    Case "Prior"
        strWHERE = "((TR_Date) <= '" & strDate1 & "')"
    Case "Since"
        strWHERE = "((TR_Date) >= '" & strDate1 & "')"
    End Select
    strWHERE = strWHERE & " AND ((TR_Status) = 3) AND ((POStatus) = 3) AND (((POL_Fulfilled) <> 'CAN') OR ((POL_Fulfilled) Is Null))" _
            & " AND (((COL_Fulfilled) <> 'CAN') OR ((COL_Fulfilled) Is Null))"
    
    strSQL = "SELECT TP_ID,TP_Name,Count(COL_ID) as QtyTitles,Sum(COL_Qty) as QtyBooks,Sum(COL_Price*COL_Qty) " _
            & "as Val, sum((COL_Price*COL_Qty)*(1-(COL_DiscountPercent/100))) as DiscountedVal,POrder, TR_Date" _
            & " FROM ReportCustomerOrders WHERE (((TP_ID) = " & lngTPID & ") AND " & strWHERE & ") " _
            & "GROUP BY TP_ID, TP_Name,POrder,TR_Date ORDER BY TR_Date"
    Set rsSub = New ADODB.Recordset
    rsSub.Open strSQL, oPC.CO
    SubRpt.object.Component rsSub
End Sub

Private Sub ghCustomer_Format()
    If rs.EOF Then Exit Sub
    
    ghCustomer.GroupValue = rs!TP_Name
End Sub

Private Sub PageFooter_Format()
    lblDate.Caption = Format(Date, "dddd, dd mmm yyyy")
End Sub

Private Sub ReportFooter_Format()
Dim rsTotals As New ADODB.Recordset
Dim strSQL As String
    
    strSQL = "SELECT Count(COL_ID) as TotalTitles, Sum(COL_Qty) as TotalBooks, " _
            & "Sum((COL_Price*COL_Qty)*(1-(COL_DiscountPercent/100))) as TotalVal FROM ReportCustomerOrders " _
            & "WHERE (" & strWHERE & ") "
    Set rsTotals = New ADODB.Recordset
    rsTotals.Open strSQL, oPC.CO
    txtReportFooter.Text = strRptHeader
    txtTotalBooks.Text = Format(rsTotals!TotalBooks, "# ##0")
    txtTotalTitles.Text = Format(rsTotals!TotalTitles, "# ##0")
    txtTotalExclVAT.Text = Format(((rsTotals!TotalVal / 100) / (1 + (oPC.Configuration.VATRate / 100))), "Currency")
    txtTotaltInclVAT.Text = Format((rsTotals!TotalVal / 100), "Currency")
    
    rsTotals.Close
    Set rsTotals = Nothing
End Sub
