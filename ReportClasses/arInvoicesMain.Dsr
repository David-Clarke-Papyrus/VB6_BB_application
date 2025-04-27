VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} arInvoicesMain 
   Caption         =   "Invoices"
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12990
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   22913
   _ExtentY        =   17701
   SectionData     =   "arInvoicesMain.dsx":0000
End
Attribute VB_Name = "arInvoicesMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As ADODB.Recordset
Dim rsSub As ADODB.Recordset
Dim WithEvents oUtil As z_UTIL
Attribute oUtil.VB_VarHelpID = -1

Dim lngTPID As Long

Dim curQtyBooks As Currency
Dim curQtyTitles As Currency
Dim curVATExcl As Currency
Dim curVATIncl As Currency

Dim strType As String
Dim strWHERE As String
Dim strDate1 As String
Dim strDate2 As String

Event Status(strMsg As String)

Public Sub Component(pRs As ADODB.Recordset, ByVal pRptHeader As String, ByVal pType As String, ByVal pDate1 As String, _
                                    ByVal pDate2 As String)
    Set rs = pRs
    txtReportHeader.text = pRptHeader
    strType = pType
    strDate1 = pDate1
    If pDate2 > "" Then strDate2 = pDate2
End Sub

Private Sub ActiveReport_ReportStart()
    Me.Left = 1000
    Me.top = 500
    Me.Height = 7000
    Me.Width = 10000
    
    Set oUtil = New z_UTIL
    
    ghCustomer.GroupValue = rs!TP_Name
    PageHeader.Height = 0
    
End Sub

Private Sub ActiveReport_Terminate()
    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    Set oUtil = Nothing
    Set SubRpt.Object = Nothing
End Sub

Private Sub Detail_AfterPrint()
    If rs.EOF Then Exit Sub
    
    ghCustomer.GroupValue = rs!TP_Name
End Sub

Private Sub Detail_Format()
    If rs.EOF Then Exit Sub
    
    txtCustomer.text = rs!TP_Name
    txtQtyBooks.text = Format(rs!QtyBooks, "# ##0")
    txtQtyTitles.text = Format(rs!QtyTitles, "# ##0")
    txtValInclVAT.text = Format((rs!DiscountedVal / 100), "Currency")
    txtValExVAT.text = Format(((rs!DiscountedVal / 100) / (1 + (oPC.Configuration.VATRate / 100))), "Currency")
    
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
    If rs.EOF Then Exit Sub
    
    Set SubRpt.Object = New arInvoicesSub
    
    strWHERE = ""
    Select Case strType
    Case "Between"
        strWHERE = "(((TR_Date) > '" & strDate1 & "') AND ((TR_Date) <= '" & strDate2 & "'))"
    Case "Prior"
        strWHERE = "((TR_Date) < '" & strDate1 & "')"
    Case "Since"
        strWHERE = "((TR_Date) > '" & strDate1 & "')"
    End Select
    strWHERE = strWHERE & " AND ((TR_Status) = 3) AND (((POL_Fulfilled) <> 'CAN') " _
            & "OR ((POL_Fulfilled) Is Null))"
    
    strSQL = "SELECT TP_Name, POrderNum, COUNT(IL_ID) as QtyTitles, SUM(IL_Qty) as QtyBooks, " _
            & " SUM((IL_Price * IL_Qty) * (1 - IL_DiscountRate)) as DiscountedVal, TR_Date " _
            & " FROM ReportInvoices WHERE (((TP_ID) = " & lngTPID & ") AND " _
            & strWHERE & ") GROUP BY TP_Name, POrderNum, TR_Date ORDER BY POrderNum"
    Set rsSub = New ADODB.Recordset
    rsSub.Open strSQL, oPC.CO
    SubRpt.Object.Component rsSub
    
End Sub

Private Sub ghCustomer_Format()
    If rs.EOF Then Exit Sub
    
    ghCustomer.GroupValue = rs!TP_Name
End Sub

Private Sub oSQL_Status(MSG As String)
    RaiseEvent Status(MSG)
End Sub

Private Sub PageFooter_Format()
    lblDate.Caption = Format(Date, "dddd, dd mmm yyyy")
End Sub

Private Sub ReportFooter_Format()
Dim rsTotals As New ADODB.Recordset
Dim strSQL As String
    strSQL = "SELECT COUNT(IL_ID) as TotalTitles, SUM(IL_Qty) as TotalBooks, " _
            & "SUM((IL_Price * IL_Qty) * (1 - IL_DiscountRate)) As TotalDiscountedVal " _
            & "From ReportInvoices WHERE (" & strWHERE & ")"
    Set rsTotals = New ADODB.Recordset
    rsTotals.Open strSQL, oPC.CO
        
    txtReportFooter.text = txtReportHeader.text
    txtTotalBooks.text = Format(rsTotals!TotalBooks, "# ##0")
    txtTotalTitles.text = Format(rsTotals!TotalTitles, "# ##0")
    txtTotalExclVAT.text = Format(((rsTotals!TotalDiscountedVal / 100) / (1 + oPC.Configuration.VATRate)), "Currency")
    txtTotaltInclVAT.text = Format((rsTotals!TotalDiscountedVal / 100), "Currency")
    
    rsTotals.Close
    Set rsTotals = Nothing
End Sub


