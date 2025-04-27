VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPOSSummary 
   BackColor       =   &H00D3D3CB&
   Caption         =   "POS summary"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   12975
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdToPDF 
      BackColor       =   &H00D5D5C1&
      Caption         =   "PDF"
      Height          =   360
      Left            =   10005
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   360
      Width           =   1380
   End
   Begin VB.CommandButton cmdToExcel 
      BackColor       =   &H00D5D5C1&
      Caption         =   "Spreadsheet"
      Height          =   360
      Left            =   11445
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1380
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      Picture         =   "frmPOSSummary.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   75
      Width           =   1000
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   7995
      Picture         =   "frmPOSSummary.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Width           =   1000
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   150
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   193658883
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4695
      TabIndex        =   2
      Top             =   135
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy HH:mm"
      Format          =   193658883
      CurrentDate     =   37421
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arPOSSummary 
      Height          =   6615
      Left            =   195
      TabIndex        =   6
      Top             =   810
      Width           =   12630
      _ExtentX        =   22278
      _ExtentY        =   11668
      SectionData     =   "frmPOSSummary.frx":0714
   End
   Begin VB.Label lblAnd 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   4275
      TabIndex        =   3
      Top             =   195
      Width           =   555
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Select period between"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   270
      TabIndex        =   0
      Top             =   180
      Width           =   3915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPOSSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bPreview As Boolean
Dim dteFrom As Date
Dim dteTo As Date
Dim bCancel As Boolean
Public enPrevPrintCSV As enumReportPresentation
Public ar As arCashupExt_ForReportsApp
Dim rs As ADODB.Recordset
Dim strStart As String
Dim strEnd As String
Dim arVoucherValue() As Long
Dim arVoucherLabel() As String
Dim lngTotalPayments As Long
Dim lngTotalVouchersINCL As Long
Dim lngTotalAccountCredits As Long
Dim oCU As New z_CashupEx


Dim dteStart As Date
Dim dteEnd As Date


Public Property Get ReportPresentation() As enumReportPresentation
    ReportPresentation = enPrevPrintCSV
End Property
Private Sub cmdClose_Click()
    bCancel = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
Dim dteLimitToView As Date
Dim oSQL As z_SQL
    If oPC.BlindCashup Then
        Set oSQL = New z_SQL
        dteLimitToView = oSQL.GetDateOfEarliestUnSignedSessionSince(dtpFrom)
        If Me.dtpTo >= StartOfDay(dteLimitToView) Then
            MsgBox "There are unsigned cash ups starting prior to your selected end date (" & Format(dteLimitToView, "dd/mm/yyyy") & "). You cannot include thse in the report. Select an earlier end date.", vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
    On Error GoTo errHandler
    oCU.Component Me.dtpFrom, Me.dtpTo
    
    dteStart = Me.dtpFrom
    dteEnd = Me.dtpTo  'DateAdd("d", 1, pENdDate)
    strStart = Year(dteStart) & "-" & Month(dteStart) & "-" & Day(dteStart)
    strEnd = Year(dteEnd) & "-" & Month(dteEnd) & "-" & Day(dteEnd)
    
    Set rs = oCU.Calculate
    
    Set ar = Nothing
    Set ar = New arCashupExt_ForReportsApp
    ar.Visible = False
    PrintCashup rs
    arPOSSummary.ReportSource = ar

EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrorIn "frmPeriodDialogue.cmdOK_Click"
End Sub
Public Function PrintCashup(pRs As ADODB.Recordset)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim i As Integer
Dim strLabel As String
Dim strValues As String
Dim lngCount As Long
Dim OpenResult As Integer
Dim lngTotalVouchersRedeemed As Long

    Set rs = pRs
    ar.Printer.Orientation = ddOPortrait
    ar.Width = 12000
    ar.Height = 8000
    ar.lblHeading = "Cash-up totals for period " & Format(dteStart, "dd/mm/yyyy") & "  to  " & Format(dteEnd, "dd/mm/yyyy")
    
    ar.lblPrinted = "Printed: " & Format(Now(), "dd/mm/yyyy HH:NN")
    ar.fCash = oCU.TotalCashInDrawerF
    ar.fCheques = oCU.TotalChequesF
    ar.fDailySalesIncl = oCU.SalesF
    lngTotalVouchersRedeemed = 0
    Do While Not rs.EOF
        i = i + 1
        ReDim Preserve arVoucherValue(i)
        ReDim Preserve arVoucherLabel(i)
        arVoucherValue(i) = FNN(rs.Fields(0))
        lngTotalVouchersRedeemed = lngTotalVouchersRedeemed + arVoucherValue(i)
        arVoucherLabel(i) = FNS(rs.Fields(1))
        strLabel = strLabel & IIf(Len(strLabel) > 0, vbCrLf, "") & arVoucherLabel(i)
        strValues = strValues & IIf(Len(strValues) > 0, vbCrLf, "") & Format(arVoucherValue(i) / oPC.Configuration.DefaultCurrency.Divisor, oPC.Configuration.DefaultCurrency.FormatString)
        rs.MoveNext
    Loop
    rs.Close


 '   ar.Component FetchPettyCash, TotalPettyCashCreditsF, TotalPettyCashWithdrawalsF, TotalPettyCashNettF
    ar.fVouchers = strValues
    ar.lblVoucher = strLabel
    ar.fWWVouchersSold = oCU.TotalVouchersSoldF
    ar.fCNIssued = oCU.TotalCNIssuedF   'Vouchers issued either as refunds or as change??????
    ar.fCNRedeemed = oCU.TotalCNRedeemedF
    ar.fDepReceived = oCU.TotalDepositsReceivedF
    ar.fDepRedeemed = oCU.TotalDepositsRedeemedF
    ar.fACRec = oCU.TotalAccountsReceivableF
    ar.fAccountReceipts = oCU.TotalAccountsPaidF
    ar.fAccountCreditnotes = oCU.TotalAccountCreditsF
    ar.fCreditCards = oCU.TotalCreditCardsF
    ar.fSalesIncl = oCU.TotalSalesInclF
    If oPC.getProperty("Cashup_Extended") = True Then
        ar.fSalesOnAccount.Text = oCU.TotalSalesOnAccountF
        ar.fSalesAll.Text = oCU.TotalSalesAllF
    End If
    
    ar.fCreditCardsRefunded = oCU.TotalCreditCardsRefundsF
    ar.fCreditCardsNett = oCU.TotalCreditCardsNettF
    ar.fTotalDiscount = oCU.TotalDiscountInclF
    ar.fLVouchersIssued = oCU.TotalLoyaltyVouchersValueF
    ar.fLoyaltyVouchersQty = "(" & oCU.TotalLoyaltyQty & ")"
    ar.fDepositsRefunded = oCU.TotalDepositsrefundedF
    ar.fDirectDeposits = oCU.TotalDirectDepositsF
    ar.fTotalPayments = oCU.TotalPaymentsF
    ar.fPCFromTill = oCU.TotalPettyCashNettF
    ar.fCOGS.Text = oCU.TotalCOGSF
   ' ar.Show vbModal
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    
    Exit Function
    
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "z_CashupEx.PrintCashup(pType)"
End Function

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    dtpFrom.Value = DateAdd("w", -1, Date)
    dtpTo.Value = Date
End Sub


'Public Property Get FromDate() As Date
'    FromDate = dteFrom
'End Property
'Public Property Get ToDate() As Date
'    ToDate = dteTo
'End Property
'Public Property Get CancelReport() As Boolean
'    CancelReport = bCancel
'End Property
'Public Property Get Preview() As Boolean
'    Preview = bPreview
'End Property
'Public Property Get NP() As Boolean
'    NP = (chkNP = 1)
'End Property
Private Sub cmdToPDF_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim pdfExpt As ActiveReportsPDFExport.ARExportPDF
    If ar Is Nothing Then Exit Sub
    ar.Run
    Set pdfExpt = New ActiveReportsPDFExport.ARExportPDF
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".PDF"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    pdfExpt.FileName = fn
    Call pdfExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enPDF
End Sub

Private Sub cmdToExcel_Click()
Dim fs As New FileSystemObject
Dim fn As String
Dim ExcelExpt As ActiveReportsExcelExport.ARExportExcel
    If ar Is Nothing Then Exit Sub
    ar.Run
    Set ExcelExpt = New ActiveReportsExcelExport.ARExportExcel
    If Not fs.FolderExists(oPC.LocalFolder & "\TEMP") Then
        fs.CreateFolder (oPC.LocalFolder & "\TEMP")
    End If
    fn = oPC.LocalFolder & "\TEMP\" & "StockValue" & Format(Now(), "YYYYMMDDHHNN") & ".XLS"
        If TryToDeleteFile(fn) = False Then
            Exit Sub
        End If
    ExcelExpt.FileName = fn
    Call ExcelExpt.Export(ar.Pages)
    OpenFileWithApplication fn, enExcel
End Sub


Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    arPOSSummary.Width = Me.Width - 600
    lngDiff = arPOSSummary.Height
    arPOSSummary.Height = Me.Height - 1600
    lngDiff = arPOSSummary.Height - lngDiff
    cmdToExcel.left = arPOSSummary.left + arPOSSummary.Width - cmdToExcel.Width
    cmdToPDF.left = arPOSSummary.left + arPOSSummary.Width - cmdToExcel.Width - cmdToPDF.Width - 10
End Sub

