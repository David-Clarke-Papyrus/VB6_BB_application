VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "z_reports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim fs As New FileSystemObject
Dim sFile As String
Event Status(strMsg As String)
Dim strMMsg As String
Attribute strMMsg.VB_VarHelpID = -1
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Public Function SalesSummary3(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date, Optional pTPCID As Long, Optional pTPSID As Long, Optional pPagebreak As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "vSalesSummary3"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@TPCID", adInteger, adParamInput, , pTPCID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@TPSID", adInteger, adParamInput, , pTPSID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set rs = cmd.execute
    Set cmd = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesSummary3(rs,pdteFrom,pdteTo)", Array(rs, pdteFrom, pdteTo)
End Function


Public Function SalesSummary2(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date, Optional pTPCID As Long, Optional pTPSID As Long, Optional pPagebreak As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "vSalesSummary2"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@TPCID", adInteger, adParamInput, , pTPCID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@TPSID", adInteger, adParamInput, , pTPSID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set rs = cmd.execute
    Set cmd = Nothing
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesSummary2(rs,pdteFrom,pdteTo)", Array(rs, pdteFrom, pdteTo)
End Function

Public Function SalesSummary1(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date, Optional pTPCID As Long, Optional pTPSID As Long, Optional pPagebreak As Boolean)
    On Error GoTo errHandler
    
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "vSalesSummary1"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@TPCID", adInteger, adParamInput, , pTPCID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@TPSID", adInteger, adParamInput, , pTPSID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set rs = cmd.execute
  '  Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesSummary1(rs,pdteFrom,pdteTo,pPagebreak)", Array(rs, pdteFrom, pdteTo, _
         pPagebreak)
End Function
Private Sub MaketSR(pdteFrom As Date, pdteTo As Date)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.CommandTimeout = 0
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "PrepareSR"
    cmd.commandType = adCmdStoredProc
    cmd.execute
 '  MsgBox "Pos 2"
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "LoadSR"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    cmd.execute
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.MaketSR(pdteFrom,pdteTo)", Array(pdteFrom, pdteTo)
End Sub
Public Function SalesByPeriod(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date, pWeekorMonth As String)
10        On Error GoTo errHandler
      Dim cmd As ADODB.Command
      Dim strTitle As String
      Dim OpenResult As Integer
      Dim prm As ADODB.Parameter
      '-------------------------------
20        OpenResult = oPC.OpenDBSHort
      '-------------------------------

30        MaketSR pdteFrom, pdteTo
40        Set cmd = New ADODB.Command
50        cmd.ActiveConnection = oPC.COShort
60        cmd.commandType = adCmdStoredProc
70        cmd.CommandTimeout = 0
80        If UCase(pWeekorMonth) = "W" Then
90            cmd.CommandText = "spvSalesByWeek"
100       Else
110           cmd.CommandText = "spvSalesByMonth"
120       End If
130       Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
140       cmd.Parameters.Append prm
150       Set prm = Nothing
160       Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
170       cmd.Parameters.Append prm
180       Set prm = Nothing
190       rs.Open cmd
200       Set cmd = Nothing
      '    rpt.Component strTitle, rs, pdteFrom, pdteTo
      '    PresentReport rpt, pReportPresentation, "SalesByPeriod", 1000, 200, 12000
      '---------------------------------------------------
      '    If OpenResult = 0 Then oPC.DisconnectDBShort
      '---------------------------------------------------

      'errHandler:
      '    If ErrMustStop Then Debug.Assert False: Resume
      '    ErrorIn "z_reports.SalesByPeriod(pdteFrom,pdteTo,pPreview,pWeekorMonth)", Array(pdteFrom, pdteTo, _
      '         pReportPresentation, pWeekorMonth)
210       Exit Function
errHandler:
220       If ErrMustStop Then Debug.Assert False: Resume
230       ErrorIn "z_reports.SalesByPeriod(rs,pdteFrom,pdteTo,pWeekorMonth)", Array(rs, pdteFrom, pdteTo, _
               pWeekorMonth)
End Function

Public Function SalesDetail(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date, bWithAppros As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim OpenResult As Integer
Dim prm As ADODB.Parameter
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
'    MaketSR pdteFrom, pdteTo
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    If bWithAppros Then
        cmd.CommandText = "vSalesDetailWithAppros"
    Else
        cmd.CommandText = "vSalesDetail"
    End If
    cmd.commandType = adCmdStoredProc
    rs.CursorLocation = adUseClient
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set rs = cmd.execute
    Set cmd = Nothing
'    rpt.Component rs, pdteFrom, pdteTo
'    PresentReport rpt, pReportPresentation, "SalesDetail", 1000, 200, 12000
'---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.SalesDetail(pdteFrom,pdteTo,pPreview)", Array(pdteFrom, pdteTo, pReportPresentation)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesDetail(rs,pdteFrom,pdteTo,bWithAppros)", Array(rs, pdteFrom, pdteTo, _
         bWithAppros)
End Function
'Function SalesSummaryReport(pdteFrom As Date, pdteTo As Date, Optional pReportPresentation As enumReportPresentation)
'    On Error GoTo errHandler
'Dim cmd As ADODB.Command
'Dim prm As ADODB.Parameter
'
'Dim dblDiscount As Double
'Dim strDiscount As String
'Dim dblSales As Double
'Dim strCashSales As String
'Dim lngQtySold As Long
'Dim strQtySold As String
'Dim dblTax As Double
'Dim strTax As String
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'    Set cmd = Nothing
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "spcud_GetTotalSales"
'    cmd.CommandType = adCmdStoredProc
'    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@Total", adInteger, adParamOutput)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@Totalexcl", adInteger, adParamOutput)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@TotalQty", adInteger, adParamOutput)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    cmd.Execute
'    dblSales = FNDBL(cmd.Parameters(2).Value) / oPC.Configuration.DefaultCurrency.Divisor
'    strCashSales = Format(dblSales, oPC.Configuration.DefaultCurrency.FormatString)
'    lngQtySold = FNN(cmd.Parameters(4).Value)
'    strQtySold = CStr(lngQtySold)
'
'    Set cmd = Nothing
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "spcud_GetTotalDiscount"
'    cmd.CommandType = adCmdStoredProc
'    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@TOTALVALUEINCL", adInteger, adParamOutput)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    cmd.Execute
'    dblDiscount = FNDBL(cmd.Parameters(2).Value)
'    strDiscount = Format(dblDiscount / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
'    rpt.Component pdteFrom, pdteTo, strCashSales, strQtySold, strDiscount, strTax
'    PresentReport rpt, pReportPresentation, "SalesSummaryReport", 1000, 200, 12000
''---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.SalesSummaryReport(pdteFrom,pdteTo,pPreview)", Array(pdteFrom, pdteTo, _
'         pReportPresentation)
'End Function

Public Function COI(rs As ADODB.Recordset, pExVAT As Boolean, pUsesLPD As Boolean, Optional ByWhat As String = "PT")
    On Error GoTo errHandler
Dim strNote As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    rs.CursorLocation = adUseClient
    If ByWhat = "P" Then
            rs.Open "SELECT Publisher as BD,SUM(QTYOH) as QTYOH,SUM(OHCOST) as OHCOST,SUM(OHLPD) as OHLPD,SUM(OHSP) as OHSP," _
                    & "SUM(APP_OS_QTY) as APP_OS_QTY,SUM(APP_OS_COST) as APP_OS_COST,SUM(APP_OS_LPD) as APP_OS_LPD,SUM(APP_OS_SP) as APP_OS_SP, " _
                    & "vCOI2_ExVAT.*,SUM(OHGM) as OHGM,SUM(OHGM_LDP) as OHGM_LDP,SUM(APP_OS_GM_COST) as APP_OS_GM_COST,SUM(APP_OS_GM_LDP) as APP_OS_GM_LDP FROM vCOI_ExVAT, vCOI2_ExVAT  " _
                    & "GROUP BY Publisher,GTOHGM,GTOHGM_LDP,GTOOGM,GTOOGM_LDP,GTMTDGM,GTMTDGM_LDP,GTAPPGM,GTAPPGM_LDP " _
                    & "ORDER BY Publisher", oPC.COShort, adOpenStatic, adLockOptimistic
    Else
        If ByWhat = "PT" Then
            rs.Open "SELECT PRODUCTTYPE as BD,SUM(QTYOH) as QTYOH,SUM(OHCOST) as OHCOST,SUM(OHLPD) as OHLPD,SUM(OHSP) as OHSP," _
                    & "SUM(APP_OS_QTY) as APP_OS_QTY,SUM(APP_OS_COST) as APP_OS_COST,SUM(APP_OS_LPD) as APP_OS_LPD,SUM(APP_OS_SP) as APP_OS_SP, " _
                    & "vCOI2_ExVAT.*,SUM(OHGM) as OHGM,SUM(OHGM_LDP) as OHGM_LDP,SUM(APP_OS_GM_COST) as APP_OS_GM_COST,SUM(APP_OS_GM_LDP) as APP_OS_GM_LDP FROM vCOI_ExVAT, vCOI2_ExVAT  " _
                    & "GROUP BY PRODUCTTYPE,GTOHGM,GTOHGM_LDP,GTOOGM,GTOOGM_LDP,GTMTDGM,GTMTDGM_LDP,GTAPPGM,GTAPPGM_LDP " _
                    & "ORDER BY PRODUCTTYPE", oPC.COShort, adOpenStatic, adLockOptimistic
        Else
            If ByWhat = "C" Then
            rs.Open "SELECT MasterCategory as BD,SUM(QTYOH) as QTYOH,SUM(OHCOST) as OHCOST,SUM(OHLPD) as OHLPD,SUM(OHSP) as OHSP," _
                    & "SUM(APP_OS_QTY) as APP_OS_QTY,SUM(APP_OS_COST) as APP_OS_COST,SUM(APP_OS_LPD) as APP_OS_LPD,SUM(APP_OS_SP) as APP_OS_SP, " _
                    & "vCOI2_ExVAT.*,SUM(OHGM) as OHGM,SUM(OHGM_LDP) as OHGM_LDP,SUM(APP_OS_GM_COST) as APP_OS_GM_COST,SUM(APP_OS_GM_LDP) as APP_OS_GM_LDP FROM vCOI_ExVAT, vCOI2_ExVAT  " _
                    & "GROUP BY MasterCategory,GTOHGM,GTOHGM_LDP,GTOOGM,GTOOGM_LDP,GTMTDGM,GTMTDGM_LDP,GTAPPGM,GTAPPGM_LDP " _
                    & "ORDER BY MasterCategory", oPC.COShort, adOpenStatic, adLockOptimistic
            End If
        End If
    End If
    strNote = "Note: Selling and cost prices are shown Ex V.A.T."
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.COI"
End Function
Public Function COI_Stocktake_Adj(pSTID As Long, pDate As String, pExVAT As Boolean, pReportPresentation As enumReportPresentation, pUsesLPD As Boolean)
    On Error GoTo errHandler
Dim ar As New arCOI_STOCKTAKE_ADJ
Dim rs As New ADODB.Recordset
Dim strNote As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    rs.CursorLocation = adUseClient
    If pExVAT Then
        rs.Open "SELECT * FROM vCOI_STOCKTAKE_ADJ_EXVAT a, vCOI2_STOCKTAKE_ADJ_EXVAT b  WHERE a.STID = " & pSTID & " AND b.STID = " & pSTID & " ORDER BY PRODUCTTYPE", oPC.COShort, adOpenStatic, adLockOptimistic
        rs.CursorLocation = adUseClient
        strNote = "Note: list and cost prices are shown Ex V.A.T."
    Else
        rs.Open "SELECT * FROM vCOI_STOCKTAKE_ADJ a, vCOI2_STOCKTAKE_ADJ b  WHERE a.STID = " & pSTID & " AND b.STID = " & pSTID & " ORDER BY PRODUCTTYPE", oPC.COShort, adOpenStatic, adLockOptimistic
        rs.CursorLocation = adUseClient
        strNote = "Note: list and cost prices are shown inclusive of V.A.T."
    End If
    ar.Component rs, pDate, strNote, pUsesLPD
    PresentReport ar, pReportPresentation, "Cost on inventory", 1000, 200, 12000
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.COI_Stocktake_Adj(pSTID,pDate)", Array(pSTID, pDate)
End Function

Function Invoices_CreditNotes(pdteFrom As Date, pdteTo As Date, Optional pReportPresentation As enumReportPresentation)
    On Error GoTo errHandler

Dim cmd As ADODB.Command
Dim rsINV As ADODB.Recordset
Dim rsCN As ADODB.Recordset
Dim rpt As New arInvoices_CNs
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

   ' pdteTo = EndOfDay(pdteTo)
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "vGetInvoiceSummary"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
  '  rsINV.CursorLocation = adUseClient
    Set rsINV = cmd.execute
  '  Set rsINV.ActiveConnection = Nothing
    Set cmd = Nothing
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "vGetCNoteSummary"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
 '   rsCN.CursorLocation = adUseClient
    Set rsCN = cmd.execute
'    Set rsCN.ActiveConnection = Nothing
    Set cmd = Nothing
    
    
    
    
    rpt.Component rsINV, rsCN, "Invoices and Credit Notes issued between " & Format(pdteFrom, "dd/mm/yyyy") & " and " & Format(pdteTo, "dd/mm/yyyy")
    PresentReport rpt, pReportPresentation, "Invoices_CreditNotes", 1000, 200, 12000

'---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.Invoices_CreditNotes(pdteFrom,pdteTo,pPreview)", Array(pdteFrom, pdteTo, _
         pReportPresentation)
End Function

Function Commissions(pStaffID As Long, pdteFrom As Date, pdteTo As Date, Optional pReportPresentation As enumReportPresentation)
    On Error GoTo errHandler

Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim rpt As New arsalesCommission
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SalesCommissions"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@StaffID", adInteger, adParamInput, , pStaffID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    rs.CursorLocation = adUseClient
    Set rs = cmd.execute
'    Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    
    rpt.Component rs, "Sales commissions between " & Format(pdteFrom, "dd/mm/yyyy") & " and " & Format(pdteTo, "dd/mm/yyyy")
    rpt.Printer.Orientation = ddOLandscape
    PresentReport rpt, pReportPresentation, "SalesCommissions", 1000, 200, 12000
'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------



    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.Commissions(pStaffID,pdteFrom,pdteTo,pPreview)", Array(pStaffID, pdteFrom, _
         pdteTo, pReportPresentation)
End Function




Function StockTakeAdjustments(Optional pReportPresentation As enumReportPresentation)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim rpt As New arStocktakeAdjustments
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_StockTakeAdjustments"
    cmd.commandType = adCmdStoredProc
    
    rs.CursorLocation = adUseClient
    Set rs = cmd.execute
 '   Set rs.ActiveConnection = Nothing
    
    rpt.Component rs
    PresentReport rpt, pReportPresentation, "StockTakeAdjustments", 1000, 200, 12000
'---------------------------------------------------
 '   If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.StockTakeAdjustments(pReportPresentation)", pReportPresentation
End Function




Function StockMovements(pdteFrom As Date, pdteTo As Date, Optional pReportPresentation As enumReportPresentation)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim rpt As arMovementsBetween
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_StockMovementsByPeriod"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    Set rs = cmd.execute
    Set rpt = New arMovementsBetween
    rpt.Component rs, pdteFrom, pdteTo
    PresentReport rpt, pReportPresentation, "StockMovements", 1000, 200, 12000
'---------------------------------------------------
 '   If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.StockMovements(pdteFrom,pdteTo,pReportPresentation)", Array(pdteFrom, pdteTo, pReportPresentation)
End Function
Function SalesByPTByDate(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "vSalesSummary5"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
   
    rs.Open cmd

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesByPTByDate(rs,pdteFrom,pdteTo)", Array(rs, pdteFrom, pdteTo)
End Function
Function SalesBySectionByDate(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date, lngcatID As Long, Optional IsSummary As Boolean)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandTimeout = 0
    
    If IsMissing(IsSummary) Then
        IsSummary = False
    End If
    
    If IsSummary = True Then
        cmd.CommandText = "vSalesSummary4summ"
    Else
        cmd.CommandText = "vSalesSummary4"
    End If
    
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing

    Set prm = cmd.CreateParameter("@CATID", adInteger, adParamInput, , lngcatID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.Open cmd
    Set cmd = Nothing
       
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesBySectionByDate(rs,pdteFrom,pdteTo)", Array(rs, pdteFrom, pdteTo)
End Function

Function SalesByCustomer(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date, pType As String, Optional bUsesLPD As String)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandTimeout = 0
    cmd.CommandText = "q_SalesByCustomer"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@SALETYPE", adVarChar, adParamInput, 20, pType)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@CostSetting", adVarChar, adParamInput, 10, bUsesLPD)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesByCustomer(rs,pdteFrom,pdteTo,pType)", Array(rs, pdteFrom, pdteTo, pType)
End Function
Function SalesByCustomerBudget(rs As ADODB.Recordset, pdteFrom As Date, pdteTo As Date)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_SalesByCustomerBudget"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
  '  rs.CursorLocation = adUseClient
    Set rs = cmd.execute
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesByCustomerBudget(rs,pdteFrom,pdteTo)", Array(rs, pdteFrom, pdteTo)
End Function

Function UnorderedDeliveryLines(pdteFrom As Date, pdteTo As Date, Optional pReportPresentation As enumReportPresentation)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim rpt As New arUnorderedDeliveries
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_UnorderedDeliveryLines"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
   
    rs.CursorLocation = adUseClient
    Set rs = cmd.execute
  '  Set rs.ActiveConnection = Nothing

    rpt.Component rs, pdteFrom, pdteTo
    PresentReport rpt, pReportPresentation, "UnorderedDeliveryLines", 1000, 200, 12000
'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------


    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.UnorderedDeliveryLines(pdteFrom,pdteTo,pReportPresentation)", Array(pdteFrom, pdteTo, _
         pReportPresentation)
End Function
Function NonStandardVAT(rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandTimeout = 0
    cmd.CommandText = "q_NonStandardVAT"
    cmd.commandType = adCmdStoredProc
    
    
    rs.CursorLocation = adUseClient
    Set rs = cmd.execute
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.NonStandardVAT(rs)", rs
End Function
Function MissingPrices(pOHLimit As Long, pPriceLimit As Long, bCostPrice As Boolean, bSP As Boolean, bRRP As Boolean, lngXMonths As Long, Optional rs As ADODB.Recordset)
    On Error GoTo errHandler

Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_MissingPrices"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    Set prm = cmd.CreateParameter("@OHLimit", adInteger, adParamInput, , pOHLimit)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@PRICELIMIT", adInteger, adParamInput, , pPriceLimit)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@CostPrice", adInteger, adParamInput, , bCostPrice)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@SP", adInteger, adParamInput, , bSP)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@RRP", adInteger, adParamInput, , bRRP)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@SalesInXMonths", adInteger, adParamInput, , lngXMonths)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.MissingPrices(pOHLimit,pPriceLimit,rs)", Array(pOHLimit, pPriceLimit, rs)
End Function
'LastStockTakeList
Function LastStockTakeList(pExVAT As Boolean, rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_StockTakeList"
    cmd.commandType = adCmdStoredProc
'    Set prm = cmd.CreateParameter("@ExVAT", adBoolean, adParamInput, , IIf(pExVAT, 1, 0))
'    cmd.Parameters.Append prm
'    Set prm = Nothing
   
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd
    Set rs = cmd.execute
   ' Set rs.ActiveConnection = Nothing
    

'---------------------------------------------------
 '   If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.LastStockTakeList(pExVAT,rs)", Array(pExVAT, rs), EA_NORERAISE
End Function

Function QtyOnHandNegative(rs As ADODB.Recordset)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_QtyOnHandNegative"
    cmd.commandType = adCmdStoredProc
    
    rs.CursorLocation = adUseClient
    Set rs = cmd.execute
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.QtyOnHandNegative(rs)", rs
End Function
Function TillAudit(rs As ADODB.Recordset, pSince As Date, pPriorTo As Date)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_TillAudit"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set prm = cmd.CreateParameter("@Since", adDate, adParamInput, , pSince)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@PriorTo", adDate, adParamInput, , pPriorTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.Open cmd
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.TillAudit(rs)", rs
End Function
Function TillJournal(rs As ADODB.Recordset, pSince As Date, pPriorTo As Date)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_TillJournal"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set prm = cmd.CreateParameter("@Since", adDate, adParamInput, , pSince)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@PriorTo", adDate, adParamInput, , pPriorTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.Open cmd
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.TillJournal(rs)", rs
End Function
Function TillLostScans(rs As ADODB.Recordset, pSince As Date, pPriorTo As Date, pTIllName As String, pStatus As Integer)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
Dim Status As Integer

'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_TillLostScans"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
 '   pTIllName = "3"
    Set prm = cmd.CreateParameter("@Tillname", adVarChar, adParamInput, 50, pTIllName)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@Start", adDate, adParamInput, , pSince)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@PrEND", adDate, adParamInput, , pPriorTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@Status", adBoolean, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.Open cmd
    pStatus = cmd.Parameters(3)
    Exit Function
errHandler:
    ErrPreserve
    If InStr(1, UCase(Error), "EXPIRED") > 0 Then
        pStatus = False
        Exit Function
    Else
        If ErrMustStop Then Debug.Assert False: Resume
        ErrorIn "z_reports.TillLostScans(rs)", rs
    End If
End Function

Function AgedStock(rs As ADODB.Recordset, pSince As Date, pPriorTo As Date, pUseLPD As Boolean, Optional SupplierID)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "q_AgedStock"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 0
    
    Set prm = cmd.CreateParameter("@Since", adDate, adParamInput, , pSince)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@PriorTo", adDate, adParamInput, , pPriorTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@SupplierID", adInteger, adParamInput, , SupplierID)
    cmd.Parameters.Append prm
    Set prm = Nothing
   
    Set prm = cmd.CreateParameter("@UseLPD", adBoolean, adParamInput, , IIf(pUseLPD, 1, 0))
    cmd.Parameters.Append prm
    Set prm = Nothing
    
   ' Set rs = cmd.execute
    
    rs.Open cmd
'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.AgedStock(pPriorTo,pReportPresentation)", Array(pPriorTo)
End Function

Function ApprosIssued(rs As ADODB.Recordset, pTPID As Long, pDate1 As Date, pDate2 As Date, pTPName As String, pTitle As String, pCosting As String) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim strWHERE As String
Dim strDate1 As String
Dim strDate2 As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    strDate1 = ReverseDate(pDate1)
    If pDate2 = CDate(0) Then
        strDate2 = ""
    Else
        strDate2 = ReverseDate(pDate2)
    End If

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    If pDate2 = CDate(0) Then
        pTitle = pTitle & " prior to " & pDate1
        strWHERE = "((TR_Date) <='" & strDate1 & "')"
    Else
        pTitle = pTitle & " between " & pDate1 & " and " & pDate2
        strWHERE = "(((TR_Date) >= '" & strDate1 & "') AND ((TR_Date) <= '" & strDate2 & "'))"
    End If
    If pTPID > 0 Then
        pTitle = pTitle & " for " & pTPName & " between " & pDate1 & " and " & pDate2
        strWHERE = strWHERE & " AND ((TP_ID) = " & pTPID & ")"
    End If
'    strSQL = "SELECT TP_ID, TP_Name, TR_Date as TranDate, TR_Code, dbo.CodeF(P_CODE,P_EAN,0) as CodeF,P_Code, P_Title, P_MainAuthor,dbo.CodeF(P_CODE,P_EAN,0) + ' ' + P_Title as FullDetails, " _
'                & "Sum(APPL_Qty ) as NetQty, " _
'                & "Sum(ExtSPIncVAT) as Gross," _
'                & "SUM(ExtRetailIncVAT) As NetAmt, " _
'                & "SUM(ExtCostExVAT) As CostExVAT " _
'                & "FROM ReportAppro WHERE (" & strWHERE _
'                & ") GROUP BY TP_ID,TP_Name,TR_Date,TR_Code,P_Code,P_Title,P_MainAuthor,P_EAN " _
'                & "HAVING (SUM(APPL_Qty) > 0) " _
'                & "ORDER BY TP_Name, TR_Code, P_Title"
'
'            rs.open strSQL, oPC.COShort
    Set rs = Appros(pTPID, pDate1, pDate2, pCosting, "ISS")

            
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.ApprosIssued(rs,pTPID,pDate1,pDate2,pTPName,pTitle)", Array(rs, pTPID, pDate1, _
         pDate2, pTPName, pTitle), EA_NORERAISE
End Function
Function ApprosOS(rs As ADODB.Recordset, pTPID As Long, pDate1 As Date, pDate2 As Date, pTPName As String, pTitle As String, pCosting As String) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim strWHERE As String
Dim strTitle1 As String
Dim strDate1 As String
Dim strDate2 As String

    strDate1 = ReverseDate(pDate1)
    If pDate2 = CDate(0) Then
        strDate2 = ""
    Else
        strDate2 = ReverseDate(pDate2)
    End If

    Set rs = New ADODB.Recordset
    If pDate2 = CDate(0) Then
        strTitle1 = pTitle & " issued prior to " & pDate1
    Else
        strTitle1 = pTitle & " issued between " & pDate1 & " and " & pDate2
    End If
    If pTPID > 0 Then
        strTitle1 = strTitle1 & " for " & pTPName '''& " between " & pDate1 & " and " & pDate2
    End If
    Set rs = Appros(pTPID, pDate1, pDate2, pCosting, "OS")
    pTitle = strTitle1

EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.ApprosOS(rs,pTPID,pDate1,pDate2,pTPName,pTitle,pCosting)", Array(rs, pTPID, _
         pDate1, pDate2, pTPName, pTitle, pCosting)
End Function

Function InvoiceSales(rsINV As ADODB.Recordset, pdteFrom As Date, pdteTo As Date)
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim rsCN As ADODB.Recordset
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "vGetInvoiceDetails"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    rsINV.Open cmd
    Set cmd = Nothing
    
    
EXIT_Handler:
    RaiseEvent Status(strMMsg & InvoiceSales)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.InvoiceSales(rsINV,pdteFrom,pdteTo,blnNoRecordsReturned)", Array(rsINV, _
'         pdteFrom, pdteTo, blnNoRecordsReturned)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.InvoiceSales(rsINV,pdteFrom,pdteTo)", Array(rsINV, pdteFrom, pdteTo)
End Function


Public Function CASHSALES(ByVal pDate1 As Date, ByVal pDate2 As Date, _
                        ByRef pNoRecordsReturned As Boolean, Optional pReportPresentation As enumReportPresentation) As String
    On Error GoTo errHandler

Dim rs As ADODB.Recordset
Dim oRPT As arFrontDeskSales
Dim strSQL As String
Dim strWHERE As String
Dim strDate1 As String
Dim strDate2 As String
Dim cmd As New ADODB.Command

Dim prm As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "FrontDeskSalesByDate"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@CSIDFrom", adDate, adParamInput, , pDate1)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@CSIDTo", adDate, adParamInput, , pDate2)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@Divisor", adInteger, adParamInput, , oPC.Configuration.DefaultCurrency.Divisor)
    cmd.Parameters.Append prm
    Set prm = Nothing
  '  rs.CursorLocation = adUseClient
    Set rs = cmd.execute
  '  Set rs.ActiveConnection = Nothing
    Set oRPT = New arFrontDeskSales
    oRPT.Component2 rs, Format(pDate1, "dd/mm/yyyy"), Format(pDate2, "dd/mm/yyyy")
    PresentReport oRPT, pReportPresentation, "CashSales", 1000, 200, 12000

'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
    strMMsg = strMMsg & "Complete"
EXIT_Handler:
    RaiseEvent Status("Front desk sales complete")
    Set oRPT = Nothing
    Exit Function

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.CASHSALES(pDate1,pDate2,pNoRecordsReturned,pReportPresentation)", Array(pDate1, _
         pDate2, pNoRecordsReturned, pReportPresentation)
End Function
Public Function FrontDeskSales(pCSIDFrom As Long, pCSIDTo As Long, pCSCodeFrom As String, pCSCodeTo As String, _
                         pNoRecordsReturned As Boolean, Optional pReportPresentation As enumReportPresentation) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oRPT As arFrontDeskSales
Dim strSQL As String
Dim strWHERE As String
Dim strDate1 As String
Dim strDate2 As String
Dim cmd As New ADODB.Command

Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------


    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "FrontDeskSales"
    cmd.commandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("@CSIDFrom", adInteger, adParamInput, , pCSIDFrom)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@CSIDTo", adInteger, adParamInput, , pCSIDTo)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    Set prm = cmd.CreateParameter("@Divisor", adInteger, adParamInput, , oPC.Configuration.DefaultCurrency.Divisor)
    cmd.Parameters.Append prm
    Set prm = Nothing
  '  rs.CursorLocation = adUseClient
    Set rs = cmd.execute
  '  Set rs.ActiveConnection = Nothing
    Set oRPT = New arFrontDeskSales
    oRPT.Component rs, pCSCodeFrom, pCSCodeTo
    PresentReport oRPT, pReportPresentation, "FrontDeskSales", 1000, 200, 12000

'---------------------------------------------------
 '   If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
    strMMsg = strMMsg & "Complete"
EXIT_Handler:
    RaiseEvent Status("Front desk sales complete")
    Set oRPT = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.FrontDeskSales(pCSIDFrom,pCSIDTo,pCSCodeFrom,pCSCodeTo,pNoRecordsReturned," & _
        "pReportPresentation)", Array(pCSIDFrom, pCSIDTo, pCSCodeFrom, pCSCodeTo, pNoRecordsReturned, _
         pReportPresentation)
End Function

Public Function CustomerOrders(ByVal pDate1 As Date, ByVal pDate2 As Date, ByVal pType As String, _
                                        ByVal pOrderBy As String, ByRef pNoRecordsReturned As Boolean, Optional pReportPresentation As enumReportPresentation) As String
    On Error GoTo errHandler
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim oRpt As arCustomerOrder
'Dim strSQL As String
'Dim strWHERE As String
'Dim strOrderBy As String
'Dim strRptHeader As String
'Dim strDate1 As String
'Dim strDate2 As String
'
'    strMMsg = "Report Status:  "
'
'    strDate1 = Format(pDate1, "yyyy/mm/dd")
'    If pDate2 > CDate(0) Then strDate2 = ReverseDate(DateAdd("d", 1, pDate2))
'    Select Case pType
'    Case "Between"
'        strWHERE = "((TR_Date) >= '" & strDate1 & "') AND ((TR_Date) < '" & strDate2 & "')"
'        strRptHeader = " between " & Format(pDate1, "dd/mm/yyyy") & " and " & Format(pDate2, "dd/mm/yyyy")
'    Case "Prior"
'        strWHERE = "((TR_Date) < '" & strDate1 & "')"
'        strRptHeader = " prior to " & pDate1
'    Case "Since"
'        strWHERE = "((TR_Date) > '" & strDate1 & "')"
'        strRptHeader = " since " & pDate1
'    End Select
'
'    If pOrderBy = "S" Then
'        strOrderBy = "QtyBooks desc"
'    Else
'        strOrderBy = "TP_Name"
'    End If
'
'    strSQL = "SELECT TP_ID,TP_Name,Count(COL_ID) as QtyTitles,Sum(COL_Qty) as QtyBooks, " _
'            & "Sum(COL_Price*COL_Qty) as Val, " _
'            & "Sum((COL_Price*COL_Qty)*(1-(COL_DiscountPercent/100))) as DiscountedVal " _
'            & "FROM ReportCustomerOrders " _
'            & "WHERE (((TR_Status) = 3) AND (((COL_Fulfilled) <> 'CAN') OR ((COL_Fulfilled) Is Null)) " _
'            & "and ((POStatus) = 3) and (((POL_Fulfilled) <> 'CAN') OR ((POL_Fulfilled) Is Null)) AND " _
'            & strWHERE & ") GROUP BY TP_ID, TP_Name ORDER BY " & strOrderBy
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, oPC.CO
'    pNoRecordsReturned = False
'    If rs.EOF Then
'        pNoRecordsReturned = True
'        GoTo EXIT_Handler
'    End If
'
'    strRptHeader = "Customer Orders" & strRptHeader
'    Set oRpt = New arCustomerOrder
'    oRpt.Component rs, strRptHeader, pType, pDate1, pDate2
'    If pPrint Then
'        oRpt.PrintReport False
'    Else
'        oRpt.Show vbModal
'    End If
'
'    CustomerOrders = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & CustomerOrders)
'    Set oRpt = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.CustomerOrders(pDate1,pDate2,pPrint,pType,pOrderBy,pNoRecordsReturned)", _
'         Array(pDate1, pDate2, pPrint, pType, pOrderBy, pNoRecordsReturned)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.CustomerOrders(pDate1,pDate2,pType,pOrderBy,pNoRecordsReturned," & _
        "pReportPresentation)", Array(pDate1, pDate2, pType, pOrderBy, pNoRecordsReturned, _
         pReportPresentation)
End Function

Public Function Deliveries(ByVal pDate As Date, ByVal pTPID As Long, _
        ByRef pNoRecordsReturned As Boolean, Optional pReportPresentation As enumReportPresentation) As String
    On Error GoTo errHandler
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim oRpt As arDeliveries
'Dim strSQL As String
'Dim strDate As String
'Dim strTPIDWhere As String
'Dim strRptHeader As String
'
'    strDate = ReverseDate(DateAdd("d", 1, pDate))
'    If pTPID > 0 Then strTPIDWhere = "AND ((TP_ID) = " & pTPID & ")"
'    strSQL = "SELECT TP_ID, TP_Name, TR_Code, TR_Date, Sum((DELL_Price * DELL_QtyTotal)) as Amt, " _
'                & "Sum((DELL_Price * DELL_QtyTotal) / (1+ (DELL_Discount/100))) as DiscountedVal,DELL_Discount " _
'                & " FROM ReportDeliveries WHERE (((TR_Status) = 3) AND ((TR_Date) > '" & strDate _
'                & "')" & strTPIDWhere & ") GROUP BY TP_ID,TP_Name,TR_Code, TR_Date,DELL_Discount ORDER BY TP_Name,TR_Code"
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, oPC.CO
'    pNoRecordsReturned = False
'    If rs.EOF Then
'        pNoRecordsReturned = True
'        GoTo EXIT_Handler
'    End If
'
'    Set oRpt = New arDeliveries
'    strRptHeader = "Deliveries outstanding prior to " & pDate
'    oRpt.Component rs, strRptHeader
'    If pPrint Then
'        oRpt.PrintReport False
'    Else
'        oRpt.Show vbModal
'    End If
'
'    Deliveries = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & Deliveries)
'    Set oRpt = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.Deliveries(pDate,pTPID,pPrint,pNoRecordsReturned)", Array(pDate, pTPID, pPrint, _
'         pNoRecordsReturned)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.Deliveries(pDate,pTPID,pNoRecordsReturned,pReportPresentation)", Array(pDate, _
         pTPID, pNoRecordsReturned, pReportPresentation)
End Function
Public Function TopSales(rs As ADODB.Recordset, pDate1 As Date, pDate2 As Date, plngSupplierID As Long, _
    plngPTID As Long, plngCATID As Long, pTop As Long) As String
    On Error GoTo errHandler
Dim strSELECT As String
Dim strGroup As String
Dim strWHERE As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    If pTop = 0 Then pTop = 100
    strSELECT = "SELECT TOP " & CStr(pTop) & " Sum(QTY) * -1 AS SoldInPeriod, tProduct.P_ID, EAN as CODEF, " _
            & " DESCR as TitleAuthorF, tProduct.P_ProductType_ID, " _
            & " tProduct.P_SupplierID, tProduct.P_QtyOnHand, tProduct.P_QtyOnOrder, tProduct.P_QtyOnBackorder, " _
            & " tProduct.P_QtyTotalSold" _
            & " FROM tProduct JOIN vSalesSummary ON tProduct.P_ID = PID "
    strGroup = " GROUP BY tProduct.P_ID, tProduct.P_Code, tProduct.P_EAN, tProduct.P_Title, tProduct.P_MainAuthor, " _
            & " tProduct.P_ProductType_ID, tProduct.P_SupplierID, tProduct.P_QtyOnHand, tProduct.P_QtyOnOrder, " _
            & " tProduct.P_QtyOnBackorder, tProduct.P_QtyTotalSold,vSalesSummary.EAN,vSalesSummary.DESCR " _
            & " ORDER BY Sum(Qty) *-1 DESC "

    strWHERE = " WHERE DTE >= '" & ReverseDate(pDate1) & "' and DTE <= '" & ReverseDate(DateAdd("d", 1, pDate2)) & "'"
    If plngPTID > 0 Then
        strWHERE = strWHERE & " AND P_ProductType_ID = " & CStr(plngPTID)
    End If
    If plngCATID > 0 Then
        strWHERE = strWHERE & " AND P_ID in (SELECT dbo.tProductSection.PSEC_P_ID " _
                & " FROM dbo.tProductSection INNER JOIN dbo.tDict ON dbo.tProductSection.PSEC_SEC_ID = dbo.tDict.DICT_ID " _
                & " WHERE PSEC_Priority = 99 and DICT_ID = " & CStr(plngCATID) & ")"
    End If
    If plngSupplierID > 0 Then
        strWHERE = strWHERE & " AND P_SUPPLIERID = " & CStr(plngSupplierID)
    End If
    
    rs.Open strSELECT & strWHERE & strGroup, oPC.COShort
    
EXIT_Handler:
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.TopSales(rs,pDate1,pDate2,plngSupplierID,plngPTID)", Array(rs, pDate1, pDate2, _
         plngSupplierID, plngPTID)
End Function
'Public Function TopSales2(pdteFrom As Date, pdteTo As Date, Optional pPagebreak As Boolean, Optional pReportPresentation As enumReportPresentation)
'    On Error GoTo errHandler
'Dim cmd As ADODB.Command
'Dim rpt As New arSalesSummary2
'Dim prm As ADODB.Parameter
'Dim rs As New ADODB.Recordset
'Dim OpenResult As Integer
''-------------------------------
'    OpenResult = oPC.OpenDBSHort
''-------------------------------
'
'    Set cmd = Nothing
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.CO
'    cmd.CommandText = "PrepareSR"
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 720
'    cmd.Execute
'
'    Set cmd = Nothing
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "LoadSR"
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 720
'    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pdteFrom)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pdteTo)
'    cmd.Parameters.Append prm
'    Set prm = Nothing
'    cmd.Execute
'
'    Set cmd = Nothing
'    Set cmd = New ADODB.Command
'    cmd.ActiveConnection = oPC.COShort
'    cmd.CommandText = "vSalesSummary3"
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandTimeout = 720
'    rs.CursorLocation = adUseClient
'    Set rs = cmd.Execute
'  '  Set rs.ActiveConnection = Nothing
'    Set cmd = Nothing
'
'    rpt.Component rs, pdteFrom, pdteTo
'    If pPagebreak Then rpt.Sections(4).NewPage = ddNPAfter
'    PresentReport rpt, pReportPresentation, "SalesSummaryByProduct", 1000, 200, 12000
'
'
''---------------------------------------------------
'   'If OpenResult = 0 Then oPC.DisconnectDBShort
''---------------------------------------------------
'
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.SalesSummary3(pdteFrom,pdteTo,pPreview,pPagebreak)", Array(pdteFrom, pdteTo, _
'          pPagebreak, pReportPresentation)
'End Function


Public Function HashNumbers(ByRef pNoRecordsReturned As Boolean) As String
    On Error GoTo errHandler
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim oRpt As arHashNumbers
'Dim strSQL As String
'
'    strSQL = "SELECT P_Code, P_Title, P_MainAuthor FROM ReportProduct " _
'            & " WHERE ((Left(P_Code,1)) = '#') ORDER BY P_Title"
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, oPC.CO
'    pNoRecordsReturned = False
'    If rs.EOF Then
'        pNoRecordsReturned = True
'        GoTo EXIT_Handler
'    End If
'
'    Set oRpt = New arHashNumbers
'    oRpt.Component rs
'    oRpt.Show vbModal
'
'    HashNumbers = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & HashNumbers)
'    Set oRpt = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.HashNumbers(pNoRecordsReturned)", pNoRecordsReturned
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.HashNumbers(pNoRecordsReturned)", pNoRecordsReturned
End Function

Public Function Invoices(ByVal pDate1 As Date, ByVal pDate2 As Date, pPrint As Boolean, ByVal pType As String, _
                                ByVal pOrderBy As String, ByRef pNoRecordsReturned As Boolean) As String
    On Error GoTo errHandler
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim oRpt As arInvoicesMain
'Dim strSQL As String
'Dim strWHERE As String
'Dim strOrderBy As String
'Dim strRptHeader As String
'Dim strDate1 As String
'Dim strDate2 As String
'
'    strMMsg = "Report Status:  "
'
'    strDate1 = ReverseDate(pDate1)
'    If pDate2 > CDate(0) Then strDate2 = ReverseDate(DateAdd("d", 1, pDate2))
'    Select Case pType
'    Case "Between", "Today"
'        strWHERE = "((TR_Date > '" & strDate1 & "') AND ((TR_Date) <= '" & strDate2 & "'))"
'        strRptHeader = "Invoices generated between " & pDate1 & " and " & pDate2
'    Case "Prior"
'        strWHERE = "(TR_Date < '" & strDate1 & "')"
'        strRptHeader = "Invoices generated prior to " & pDate1
'    Case "Since"
'        strWHERE = "(TR_Date > '" & strDate1 & "')"
'        strRptHeader = "Invoices generated since " & pDate1
'    End Select
''    strWHERE = strWHERE & " AND ((TR_Status) = 3) AND ((POStatus) = 3))"
'    strWHERE = strWHERE & " AND ((TR_Status = 3) or (TR_Status = 4) )"
'
'    If pOrderBy = "S" Then
'        strOrderBy = "QtyBooks desc"
'    Else
'        strOrderBy = "TP_Name"
'    End If
'
'    strSQL = "SELECT TP_ID, TP_Name, COUNT(IL_ID) as QtyTitles, Sum(IL_Qty) as QtyBooks," _
'            & " Sum((IL_Price * IL_Qty) * (1-(IL_DiscountRate/100))) as DiscountedVal FROM ReportInvoices" _
'            & " WHERE (" & strWHERE & ") GROUP BY TP_ID, TP_Name ORDER BY " & strOrderBy
'
'    RaiseEvent Status("Processing Report . . .")
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, oPC.CO
'    If rs.EOF Then
'        pNoRecordsReturned = True
'        GoTo EXIT_Handler
'    End If
'
'    Set oRpt = New arInvoicesMain
'    oRpt.Component rs, strRptHeader, pType, strDate1, strDate2
'    If pPrint Then
'        oRpt.PrintReport False
'    Else
'        oRpt.Show vbModal
'    End If
'
'    Invoices = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & Invoices)
'    Set oRpt = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.Invoices(pDate1,pDate2,pPrint,pType,pOrderBy,pNoRecordsReturned)", _
'         Array(pDate1, pDate2, pPrint, pType, pOrderBy, pNoRecordsReturned)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.Invoices(pDate1,pDate2,pPrint,pType,pOrderBy,pNoRecordsReturned)", _
         Array(pDate1, pDate2, pPrint, pType, pOrderBy, pNoRecordsReturned)
End Function

Function PurchaseOrders(ByVal pTPID As Long, ByVal pPrint As Boolean, ByRef pNoRecordsReturned As Boolean) As String
    On Error GoTo errHandler
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim strSQL As String
'Dim strWHERE As String

    
'    strMMsg = "Report Status  "
'    Set rs = New ADODB.Recordset
'    If pTPID > 0 Then
'        strWHERE = " AND ((TP_ID) = " & pTPID & ")"
'    End If
'    strSQL = "SELECT TP_Name, IsNull(TR_Date,TR_CaptureDate) as TranDate,P_QtyOnHand, " _
'            & "TR_Code, P_Code, P_Title, P_MainAuthor, P_Publisher,POL_QtySS, POL_QtyFirm, " _
'            & "POL_Price, POL_ForeignPrice, POL_QtyReceivedSoFar,P_LastDateOrdered " _
'            & "FROM ReportPurchaseOrders WHERE ((((POL_Fulfilled) <> 'CAN') OR ((POL_FulFilled) Is Null)) " _
'            & "AND ((POL_QtySS + POL_QtyFirm) < POL_QtyReceivedSoFar) " _
'            & "AND ((POL_ETA) < '" & ReverseDate(Date) & "')" & strWHERE & ") ORDER BY TP_Name"
'    rs.Open strSQL, oPC.CO
'    pNoRecordsReturned = False
'    If rs.EOF Then
'        pNoRecordsReturned = True
'        GoTo EXIT_Handler
'     End If
'
'    Set oRpt = New arPurchaseOrdersOutstanding
'    oRpt.Component rs
'    If pPrint Then
'        oRpt.PrintReport False
'    Else
'        oRpt.Show
'    End If
'
'    PurchaseOrders = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & PurchaseOrders)
'    Set oRpt = Nothing
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.PurchaseOrders(pTPID,pPrint,pNoRecordsReturned)", Array(pTPID, pPrint, _
'         pNoRecordsReturned)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.PurchaseOrders(pTPID,pPrint,pNoRecordsReturned)", Array(pTPID, pPrint, _
         pNoRecordsReturned)
End Function

'Function GenerateReorderStock(bAll As Boolean, pTPID, pEnterSince, pDate As Date, iPrintOrPreview As PreviewPrint, lngRecordsReturned As Long, _
'    pTitle As String, pTitle1 As String, pFooter As String, bNoRecordsReturned As Boolean)
'Function ReorderStock(ByVal pWeeksback As Integer, ByVal pPrint As Boolean, ByVal bNoRecordsReturned As Boolean) As String
'On Error GoTo ERR_Handler
'
''Dim oarReorderStock As arReorderStock
'Dim rs As ADODB.Recordset
'Dim strTitle As String
'Dim strTitle1 As String
'Dim strFooter As String
'Dim CW As Integer
'Dim pAll As Boolean
'Dim strTPName As String
'Dim rsIN As ADODB.Recordset
'Dim rsOut As ADODB.Recordset
'Dim strLastOrders As String
'Dim strDate As String
'
'    Set rs = New ADODB.Recordset
'    CW = oPC.Configuration.CurrentWeek - pWeeksback
'    strMMsg = "Report Status  "
'''''''    strDate = pDate
'''''''    retval = oSQL.DropTable("TEMP_r4", "Erasing table . . . ")
'''''''    retval = oSQL.DropTable("TEMP_r4_1", "Erasing table . . . ")
'''''''    retval = oSQL.DropTable("TEMP_r4_1c", "Erasing table . . . ")
'''''''    retval = oSQL.DropTable("TEMP_r4_2a", "Erasing table . . . ")
'''''''    retval = oSQL.DropTable("TEMP_r4_2b", "Erasing table . . . ")
'''''''    retval = oSQL.DropTable("TEMP_r4_3b", "Erasing table . . . ")
'''''''    retval = oSQL.DropTable("TEMP_r4_4a", "Erasing table . . . ")
'''''''    retval = oSQL.DropTable("TEMP_r4_4b", "Erasing table . . . ")
'''''''    retval = oSQL.DropTable("Temp_r4_5", "Erasing table . . . ")
'''''''    lngRecordsReturned = oSQL.RunProc("q_r4_1a", Array(strDate, pTPID), "Processing 1/5  . . .")
'''''''    lngRecordsReturned = oSQL.RunProc("q_r4_1b", Array(pTPID, CW), "Processing 2/5  . . .")
'''''''    lngRecordsReturned = oSQL.RunProc("q_r4_1c", Array(), "Processing 3/5  . . .")
'''''''    lngRecordsReturned = oSQL.RunProc("q_r4_2a", Array(strDate), "Processing 4/5  . . .")
'''''''    lngRecordsReturned = oSQL.RunProc("q_r4_2b", Array(CW), "Processing 5/5  . . .")
'''''''
'''''''    Dim iCount As Integer
'''''''    Dim lngCurPID As Long
'''''''    Dim lngPrevPID As Long
'''''''
'''''''    Set rsIN = New ADODB.Recordset
'''''''    Set rsOut = New ADODB.Recordset
'''''''    lngRecordsReturned = oSQL.RunGetRecordset("q_r4_3a", enQuery, Array(), "Processing Report  . . . ", rsIN)
'''''''    retval = oSQL.RunProc("q_r4_3b", Array(), "Processing Report  Blank table . . .")
'''''''    rsOut.Open "TEMP_r4_3b", oPC.CO, adOpenDynamic, adLockOptimistic
'''''''
'''''''    Do While Not rsIN.EOF
'''''''        iCount = 0
'''''''        strLastOrders = ""
'''''''        lngCurPID = rsIN.Fields("PID")
'''''''        lngPrevPID = lngCurPID
'''''''        Do While lngCurPID = lngPrevPID And Not rsIN.EOF
'''''''                iCount = iCount + 1
'''''''                If iCount <= 3 Then
'''''''                    If Len(strLastOrders) > 0 Then strLastOrders = strLastOrders & vbCrLf
'''''''                    strLastOrders = strLastOrders + Format(rsIN.Fields("TR_Date"), "dd/mm/yy") & "   " & rsIN.Fields("SOL_Qty") _
'''''''                        & " / " & rsIN.Fields("SOL_SSQty")
'''''''                End If
'''''''                rsIN.MoveNext
'''''''                If Not rsIN.EOF Then lngCurPID = rsIN.Fields("PID")
'''''''        Loop
'''''''        rsOut.AddNew
'''''''        rsOut.Fields("PID") = lngPrevPID
'''''''        rsOut.Fields("LastOrders") = strLastOrders
'''''''        rsOut.Update
'''''''    Loop
'''''''    rsIN.Close
'''''''    rsOut.Close
'''''''    Set rsIN = Nothing
'''''''    Set rsOut = Nothing
'''''''
'''''''    lngRecordsReturned = oSQL.RunProc("q_r4_4a", Array(oPC.Configuration.CurrentMonth), "Processing 1/6  . . .")
'''''''    lngRecordsReturned = oSQL.RunProc("q_r4_4b", Array(), "Processing 2/6  . . .")
'''''''    lngRecordsReturned = oSQL.RunGetRecordset("q_r4", enQuery, Array(oPC.Configuration.CurrentMonth), "Processing 6/6  . . .", rs)
''''''''    lngRecordsReturned = oSQL.RunGetRecordset("TEMP_r4", Table, Array(), " Writing report . . .", rs)
'''''''
'''''''    strTitle = pTitle
'''''''    strTitle1 = pTitle1
'''''''    strTPName = ""
'''''''    strFooter = pFooter
'''''''    pAll = bAll
'    bNoRecordsReturned = False
'    If rs.EOF Then
'        bNoRecordsReturned = True
'        GoTo EXIT_Handler
'     End If
'  '  Set oarReorderStock = New arReorderStock
' '  oarReorderStock.Printer.Orientation = ddOLandscape
''    oarReorderStock.Component rs, pAll, strTitle, strTitle1, strTPName, pFooter
''    If pPrint = pPrint Then
''        oarReorderStock.PrintReport False
''    Else
''        oarReorderStock.Show
''    End If
'   ReorderStock = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & ReorderStock)
'    Exit Function
'ERR_Handler:
'    ReorderStock = Err.Description
'    Err.Clear
'    GoTo EXIT_Handler
'End Function

'Public Function Seesafe(ByVal pDate As Date, ByVal pTPID As Long, ByVal pPrint As Boolean, ByRef pNoRecordsReturned As Boolean) As String
'Dim rs As ADODB.Recordset
'Dim oRpt As arSeeSafeReturns
'Dim strSQL As String
'Dim strDate As String
'Dim strWHERE As String
'    On Error GoTo ERR_Handler
'
'    If pTPID > 0 Then strWHERE = " AND TP_ID = " & pTPID
'    strDate = ReverseDate(DateAdd("d", 1, pDate))
'    strSQL = "SELECT P_Title, P_Code, P_MainAuthor, P_QtyOnHand, Max(TR_Date) as TranDate, TP_Name,P_Publisher, " _
'            & "Sum(POL_QtySS) as QtySS FROM tTP JOIN tTR ON tTP.TP_ID = tTR.TR_TP_ID JOIN tDELL ON tTR.TR_ID = tDELL.DELL_TR_ID " _
'            & "JOIN tPOL JOIN tProduct ON tPOL.POL_P_ID = tProduct.P_ID ON tDELL.DELL_POL_ID = tPOL.POL_ID " _
'            & "WHERE (((TR_Date) < '" & strDate & "') AND ((POL_QtySS) > 0) AND ((P_QtyOnHand) > 0)" & strWHERE & ") " _
'            & "GROUP BY P_Title,P_Code,P_MainAuthor,P_QtyOnHand, " _
'            & "TP_Name,P_Publisher ORDER BY TP_Name,P_Title"
'    Set rs = New ADODB.Recordset
'    rs.Open strSQL, oPC.CO
'    pNoRecordsReturned = False
'    If rs.EOF Then
'        pNoRecordsReturned = True
'        GoTo EXIT_Handler
'    End If
'
'    Set oRpt = New arSeeSafeReturns
'    oRpt.Component rs
'    If pPrint Then
'        oRpt.PrintReport False
'    Else
'        oRpt.Show vbModal
'    End If
'    Seesafe = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & Seesafe)
'    Exit Function
'ERR_Handler:
'    Seesafe = Err.Description
'    Err.Clear
'    GoTo EXIT_Handler
'    Resume
'End Function

Public Function TradingPartners(ByVal pRoleType As Integer, ByRef pNoRecordsReturned As Boolean) As String
    On Error GoTo errHandler
'    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim oRpt As arSuppliers
'Dim strSQL As String
'
'    Set rs = New ADODB.Recordset
'    strSQL = "SELECT * FROM ReportTPs WHERE ((TP_Role) = " & pRoleType & ") ORDER BY TP_Name"
'    rs.Open strSQL, oPC.CO
'    pNoRecordsReturned = False
'    If rs.EOF Then
'        pNoRecordsReturned = True
'        GoTo EXIT_Handler
'    End If
'
'    Set oRpt = New arSuppliers
'    oRpt.Component rs
'    oRpt.Show vbModal
'
'    TradingPartners = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & TradingPartners)
'    Exit Function
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.TradingPartners(pRoleType,pNoRecordsReturned)", Array(pRoleType, _
'         pNoRecordsReturned)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.TradingPartners(pRoleType,pNoRecordsReturned)", Array(pRoleType, _
         pNoRecordsReturned)
End Function
Public Function SalesItems(ByVal pDate1 As Date, ByVal pDate2 As Date, ByVal pType As String, _
        ByVal pOrderBy As String, ByRef pNoRecordsReturned As Boolean, Optional pReportPresentation As enumReportPresentation) As String
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim oRPT As arSalesItems
Dim strSQL As String
Dim strWHERE As String
Dim strOrderBy As String
Dim strRptHeader As String
Dim strDate1 As String
Dim strDate2 As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    strMMsg = "Report Status:  "

    Select Case pType
    Case "Between", "Today"
        strWHERE = "((Dte > '" & ReverseDate(pDate1) & "') AND ((Dte) <= '" & ReverseDate(pDate2) & "'))"
        strRptHeader = "Sales between " & pDate1 & " and " & pDate2
    Case "Prior"
        strWHERE = "(Dte < '" & strDate1 & "')"
        strRptHeader = "Sales prior to " & pDate1
    Case "Since"
        strWHERE = "(Dte > '" & strDate1 & "')"
        strRptHeader = "Sales since " & pDate1
    End Select

    If pOrderBy = "S" Then
        strOrderBy = "Qty desc"
    Else
        strOrderBy = "Title"
    End If

    strSQL = "SELECT TITLE,SUBTITLE,CODE,EAN,SaleType,SP,SUM(qty) as QTY,SUM(Valu) as VALU, SUM(NetValu) as NetValu,SUM(ExVAT) as ExVat FROM vSALEItems" _
            & " WHERE (" & strWHERE & ")  " _
            & " GROUP BY TITLE,SUBTITLE,CODE,EAN,SALETYPE,SP " _
            & " ORDER BY " & strOrderBy

    RaiseEvent Status("Processing Report . . .")
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open strSQL, oPC.COShort
 '   Set rs.ActiveConnection = Nothing
    If rs.EOF Then
        pNoRecordsReturned = True
        GoTo EXIT_Handler
    End If

    Set oRPT = New arSalesItems
    oRPT.Component rs, pDate1, pDate2
    PresentReport oRPT, pReportPresentation, "SalesItems"


    SalesItems = ""
    strMMsg = strMMsg & "Complete"
'---------------------------------------------------
 '   If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
EXIT_Handler:
    RaiseEvent Status(strMMsg)
    Set oRPT = Nothing
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.SalesItems(pDate1,pDate2,pType,pOrderBy,pNoRecordsReturned," & _
        "pReportPresentation)", Array(pDate1, pDate2, pType, pOrderBy, pNoRecordsReturned, _
         pReportPresentation)
End Function

Public Function ValueOfDeliveries(ByVal pFrom As Date, ByVal pTo As Date, ByVal pTPID As Long, _
        ByRef pNoRecordsReturned As Boolean, Optional pReportPresentation As enumReportPresentation) As String
    On Error GoTo errHandler
'Dim rs As ADODB.Recordset
'Dim oRpt As arDeliveries
'Dim strDate As String
'Dim strTPIDWhere As String
'Dim strRptHeader As String
'    On Error GoTo ERR_Handler
'
'    strDate = ReverseDate(DateAdd("d", 1, pDate))
'    If pTPID > 0 Then strTPIDWhere = "AND ((TP_ID) = " & pTPID & ")"
'    Set rs = New ADODB.Recordset
'    rs.Open "SELECT *  FROM rpt_ValueOfDeliveries_1 WHERE", oPC.CO, adOpenForwardOnly
'    pNoRecordsReturned = False
'    If rs.EOF Then
'        pNoRecordsReturned = True
'        GoTo EXIT_Handler
'    End If
'
'    Set oRpt = New arDeliveries
'    strRptHeader = "Deliveries outstanding prior to " & pDate
'    oRpt.Component rs, strRptHeader
'    If pPrint Then
'        oRpt.PrintReport False
'    Else
'        oRpt.Show vbModal
'    End If
'
'    Deliveries = ""
'    strMMsg = strMMsg & "Complete"
'EXIT_Handler:
'    RaiseEvent Status(strMMsg & Deliveries)
'    Set oRpt = Nothing
'    Exit Function
'ERR_Handler:
'    Deliveries = Err.Description
'    Err.Clear
'    GoTo EXIT_Handler
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.ValueOfDeliveries(pFrom,pTO,pTPID,pNoRecordsReturned,pReportPresentation)", _
         Array(pFrom, pTo, pTPID, pNoRecordsReturned, pReportPresentation)
End Function




Public Function ConsolidatedGRNs(CGRN As Long, blnNoRecordsReturned As Boolean, Optional pReportPresentation As enumReportPresentation)
    On Error GoTo errHandler

Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim rpt As New arCGRN
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT * FROM zvCGRN WHERE DEL_CGRNNumber = " & CGRN
    cmd.commandType = adCmdText
    
    Set rs = cmd.execute
    Set cmd = Nothing
    
    If rs.EOF And rs.BOF Then
        blnNoRecordsReturned = True
        Exit Function
    Else
        blnNoRecordsReturned = False
    End If
    rpt.Component rs, "CONSOLIDATED GOODS RECEIVED NOTE: No." & CStr(CGRN) & vbCrLf & "Date: " & Format(rs!ReceivedDate, "dd-mm-yyyy") & vbCrLf & oPC.Configuration.DefaultStore.Description & "(" & oPC.Configuration.DefaultStore.code & ")" & vbCrLf & vbCrLf & rs.Fields("TP_NAME") & " " & rs.Fields("TP_ACNO")
    PresentReport rpt, pReportPresentation, "ConsolidatedGRNs", 1000, 200, 12000
'---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.ConsolidatedGRNs(CGRN,blnNoRecordsReturned,pReportPresentation)", Array(CGRN, _
         blnNoRecordsReturned, pReportPresentation)
End Function
Function DiscrepancyReport(pType As String, bExVat As Boolean, bPartial As Boolean, Optional pReportPresentation As enumReportPresentation)
Dim arB As arValidation_B
Dim tmpNumber As Long
Dim rs As ADODB.Recordset
Dim strSQL As String
Dim strTitle As String
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

        Set arB = New arValidation_B
        arB.Printer.Orientation = ddOLandscape
        If bExVat Then
            If pType = "ALL" Then
                If bPartial Then
                    strSQL = "Select * FROM vDiscrepancyExVAT WHERE NOT a.PID is NULL ORDER BY SEC,P_TITLE"
                Else
                    strSQL = "Select * FROM vDiscrepancyExVAT ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (All values Ex VAT)"
            ElseIf pType = "POS" Then
                If bPartial Then
                    strSQL = "Select * FROM vDiscrepancyExVAT WHERE NOT a.PID is NULL AND DIFF < 0 ORDER BY SEC,P_TITLE"
                Else
                    strSQL = "Select * FROM vDiscrepancyExVAT WHERE DIFF < 0 ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (Adjustment up)  (All values Ex VAT)"
            ElseIf pType = "NEG" Then
                If bPartial Then
                    strSQL = "Select * FROM vDiscrepancyExVAT WHERE NOT a.PID is NULL AND  DIFF > 0  ORDER BY SEC,P_TITLE"
                Else
                    strSQL = "Select * FROM vDiscrepancyExVAT WHERE DIFF > 0  ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (Adjustment down) (All values Ex VAT)"
            End If
        Else
            If pType = "ALL" Then
                If bPartial Then
                    strSQL = "Select * FROM vDiscrepancy WHERE NOT a.PID is NULL ORDER BY SEC,P_TITLE"
                Else
                    strSQL = "Select * FROM vDiscrepancy ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (All values Incl VAT)"
            ElseIf pType = "POS" Then
                If bPartial Then
                    strSQL = "Select * FROM vDiscrepancy WHERE NOT a.PID is NULL AND DIFF < 0 ORDER BY SEC,P_TITLE"
                Else
                    strSQL = "Select * FROM vDiscrepancy WHERE DIFF < 0 ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (Adjustment up) (All values Incl VAT)"
            ElseIf pType = "NEG" Then
                If bPartial Then
                    strSQL = "Select * FROM vDiscrepancy WHERE NOT a.PID is NULL AND  DIFF > 0  ORDER BY SEC,P_TITLE"
                Else
                    strSQL = "Select * FROM vDiscrepancy WHERE DIFF > 0  ORDER BY SEC,P_TITLE"
                End If
                strTitle = "Stock adjustments (Adjustment down) (All values Incl VAT)"
            End If
        End If
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        rs.Open strSQL, oPC.COShort, adOpenStatic
     '   Set rs.ActiveConnection = Nothing
        arB.Component rs, strTitle
        Screen.MousePointer = vbDefault
        PresentReport arB, pReportPresentation, strTitle, 1000, 400, 12000, 6000
        Set rs = Nothing
        Set arB = Nothing
'---------------------------------------------------
 ' If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

End Function

Function DuplicateEAN(Optional pReportPresentation As enumReportPresentation)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim rpt As New arDuplicateEANs
Dim prm As ADODB.Parameter

Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT * FROM vDuplicateEANs"
    cmd.commandType = adCmdText
    
   
    Set rs = cmd.execute
    
    rpt.Component rs
    PresentReport rpt, pReportPresentation, "DuplicateEAN", 1000, 200, 12000
'---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.DuplicateEAN(pReportPresentation)", pReportPresentation
End Function
Function DuplicateShortCodes(Optional pReportPresentation As enumReportPresentation)
    On Error GoTo errHandler
Dim cmd As New ADODB.Command
Dim rs As New ADODB.Recordset
Dim rpt As New arDuplicateShortCodes
Dim prm As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    cmd.ActiveConnection = oPC.COShort
    cmd.CommandText = "SELECT * FROM vDuplicateCodes"
    cmd.commandType = adCmdText
    
   
    Set rs = cmd.execute
    
    rpt.Component rs
    PresentReport rpt, pReportPresentation, "DuplicateCodes", 1000, 200, 12000

'---------------------------------------------------
'    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.DuplicateShortCodes(pReportPresentation)", pReportPresentation
End Function

Public Function Filenames() As ADODB.Recordset
    On Error GoTo errHandler
Dim iresult As Long
Dim cmd As ADODB.Command
Dim par As ADODB.Parameter
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = New ADODB.Command
    cmd.CommandText = "sp_STOCKTAKE_GETFILENAMES"
    cmd.commandType = adCmdStoredProc
    
    
    cmd.ActiveConnection = oPC.COShort
    Set Filenames = cmd.execute
    
    Set cmd = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.Filenames"
End Function

Public Sub PresentReport(rpt As Object, pReportPresentation As enumReportPresentation, pReportName As String, Optional T As Long, Optional L As Long, Optional w As Long, Optional H As Long, Optional pOrientation As String)
Dim strExecutable As String
Dim i As Integer
    If Not IsMissing(pOrientation) Then
        If pOrientation = "L" Then
            rpt.Printer.Orientation = ddOLandscape
        Else
            rpt.Printer.Orientation = ddOPortrait
        End If
    End If
    If pReportPresentation = enPreview And Not rpt.WindowState = 2 Then
        If T > 0 Then rpt.top = T
        If L > 0 Then rpt.Left = L
        If w > 0 Then rpt.Width = w
        If H > 0 Then rpt.Height = H
        
        rpt.Show
    ElseIf pReportPresentation = enPrintout Then
        rpt.PrintReport True
    ElseIf pReportPresentation = enCSV Then
        rpt.Run False
        If Not fs.FolderExists(oPC.LocalFolder & "TEMP") Then
            fs.CreateFolder oPC.LocalFolder & "TEMP"
        End If
        sFile = oPC.LocalFolder & "TEMP\" & pReportName & ".XLS"
        i = 0
        Do Until fs.FileExists(sFile) = False
            i = i + 1
            sFile = oPC.LocalFolder & "TEMP\" & pReportName & "_" & CStr(i) & ".XLS"
        Loop
        
        
        
        xls.FileName = sFile
        If rpt.Pages.Count > 0 Then
            xls.Export rpt.Pages
        End If
        strExecutable = GetPDFExecutable(oPC.SharedFolderRoot & "\TEMPLATES\DUMMY.XLS")
        If strExecutable = "" Then
            MsgBox "Contact support, missing 'DUMMY.XLS' file in \Templates folder or Excel not installed." & vbCrLf & "Report will not open now but is saved in " & sFile, vbInformation, "Status"
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        Shell strExecutable & " " & """" & sFile & """", vbNormalFocus
    End If
    Screen.MousePointer = vbDefault

End Sub

Public Function GetStockReconciliation(pDate As Date, bInclusive As Boolean, bPerProduct As Boolean) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetStockreconciliation"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pDate)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Inclusive", adBoolean, adParamInput, , bInclusive)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@PerProduct", adBoolean, adParamInput, , bPerProduct)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
  '  Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    
    Set GetStockReconciliation = rs
'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function
Public Function GetStockReconciliationSUM(pFromDate As Date, pToDate As Date, pPerProduct As Boolean) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetStockreconciliationSum"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pFromDate)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pToDate)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@PerProduct", adBoolean, adParamInput, , pPerProduct)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
  '  Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    
    Set GetStockReconciliationSUM = rs
'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function
Public Function GetStockReconciliationDetails(pFromDate As Date, pToDate As Date, pCode As String) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetStockreconciliationDetails"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    Set prm = cmd.CreateParameter("@FROM", adDate, adParamInput, , pFromDate)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@To", adDate, adParamInput, , pToDate)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Code", adVarChar, adParamInput, 20, pCode)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
  '  Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    
    Set GetStockReconciliationDetails = rs
'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function

Public Function GetPOStats(pDayCount As Integer) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetPOStats"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    Set prm = cmd.CreateParameter("@DaysBack", adInteger, adParamInput, , pDayCount)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
   ' Set rs.ActiveConnection = Nothing
    Set cmd = Nothing
    
    Set GetPOStats = rs
'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
End Function


Public Function Appros(pTPID As Long, pDate1 As Date, pDate2 As Date, pCosting As String, pType As String) As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    If pType = "OS" Then
        cmd.CommandText = "OSAppros"
    Else
        cmd.CommandText = "ApprosIssued"
    End If
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Date1", adDate, adParamInput, , ReverseDate(pDate1))
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Date2", adDate, adParamInput, , ReverseDate(pDate2))
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Costing", adVarChar, adParamInput, 10, pCosting)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
    Set cmd = Nothing
    
    Set Appros = rs
'---------------------------------------------------
  '  If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
End Function

Function DebtorsReport(rs As ADODB.Recordset, pNoRecordsReturned As Boolean) As String
    On Error GoTo errHandler
Dim strSQL As String
Dim strWHERE As String
Dim strMsg As String
Dim strFooter As String
Dim strTitle1 As String
Dim strDate1 As String
Dim strDate2 As String

    Set rs = New ADODB.Recordset
    strMMsg = "Report Status:  "
    RaiseEvent Status(strMMsg & "Executing report  . . .")
    Set rs = GetDebtorsReport()
    pNoRecordsReturned = False
    If rs.EOF Then
        pNoRecordsReturned = True
        GoTo EXIT_Handler
     End If
     RaiseEvent Status(strMMsg & "Processing report  . . .")
    DebtorsReport = ""
    strMMsg = strMMsg & "Complete"
EXIT_Handler:
    RaiseEvent Status(strMMsg & DebtorsReport)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "z_reports.DebtorsReport(pReportPresentation)", pReportPresentation
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.DebtorsReport(rs,pNoRecordsReturned,pReportPresentation)", Array(rs, _
         pNoRecordsReturned)
End Function

Public Function GetDebtorsReport() As ADODB.Recordset
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetDebtorsReport"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
    Set cmd = Nothing
    
    Set GetDebtorsReport = rs
End Function

'GetPerformanceData
Public Function GetPerformanceData_Supplier(pTPID As Long, pTop As Long) As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetPerformanceData_Supplier"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    Set prm = cmd.CreateParameter("@TPID", adInteger, adParamInput, , pTPID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    Set prm = cmd.CreateParameter("@Top", adInteger, adParamInput, , pTop)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
    Set cmd = Nothing
    
    Set GetPerformanceData_Supplier = rs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.GetPerformanceData_Supplier(pTPID)", pTPID
End Function

Public Function GetPerformanceData_Category(pDictID) As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetPerformanceData_Category"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    Set prm = cmd.CreateParameter("@DICTID", adInteger, adParamInput, , pDictID)
    cmd.Parameters.Append prm
    Set prm = Nothing
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
    Set cmd = Nothing
    
    Set GetPerformanceData_Category = rs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.GetPerformanceData_Category(pDictID)", pDictID
End Function
Public Function GetPerformanceData_Summary() As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetPerformanceData_Summary"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 720
    
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
    Set cmd = Nothing
    
    Set GetPerformanceData_Summary = rs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.GetPerformanceData_Summary"
End Function

Public Function GetPaymentOrder() As ADODB.Recordset
    On Error GoTo errHandler
Dim cmd As ADODB.Command
Dim prm As ADODB.Parameter
Dim rs As New ADODB.Recordset
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    Set cmd = Nothing
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = oPC.COShort
    
    cmd.CommandText = "GetPaymentOrder"
    cmd.commandType = adCmdStoredProc
    cmd.CommandTimeout = 3000
    
    
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockOptimistic
    Set cmd = Nothing
    
    Set GetPaymentOrder = rs
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_reports.GetPaymentOrder"
End Function

