VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00552619&
   Caption         =   "Papyrus II reports v1.3"
   ClientHeight    =   7830
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13455
   Icon            =   "frmPBKSMainReports.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7455
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18071
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   13455
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      Begin VB.Image imgLogo 
         Height          =   1800
         Left            =   1200
         Picture         =   "frmPBKSMainReports.frx":038A
         Top             =   -75
         Width           =   4755
      End
      Begin VB.Image imgLogoMask 
         Height          =   1800
         Left            =   7305
         Picture         =   "frmPBKSMainReports.frx":1C20E
         Top             =   120
         Width           =   4755
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3285
      Top             =   5070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":38092
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3862C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":38BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":39160
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":396FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":39C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3A22E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3A7C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3AD62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3B2FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3B896
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3BE30
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3C3CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3C964
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3CEFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3D498
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3DA32
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3DFCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3E566
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3EB00
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3F09A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3F634
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMainReports.frx":3FBCE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSwapDB 
         Caption         =   "Swap to TEST database"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuAppros 
      Caption         =   "Appros"
      Begin VB.Menu mnuApprosIssued 
         Caption         =   "A&ppros Issued"
      End
      Begin VB.Menu mnuReportsApprosOS 
         Caption         =   "&Appros outstanding"
      End
      Begin VB.Menu mnuAgedAppros 
         Caption         =   "Aged appros &outstanding (cube)"
      End
   End
   Begin VB.Menu mnuCOI 
      Caption         =   "Inventory"
      Begin VB.Menu mnuStockList 
         Caption         =   "Stock list"
      End
      Begin VB.Menu mnuCOIN 
         Caption         =   "Value of Inventory - Normal"
      End
      Begin VB.Menu mnuCOIST 
         Caption         =   "Value of Inventory-Stocktake adjustments"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAged 
         Caption         =   "A&ged stock"
      End
      Begin VB.Menu mnuStatistics 
         Caption         =   "Statistics"
         Begin VB.Menu mnuDailyValues 
            Caption         =   "Daily values"
         End
         Begin VB.Menu mnuStockRecon 
            Caption         =   "Stock reconciliation"
         End
         Begin VB.Menu mnuRerun 
            Caption         =   "Re-run movements since stocktake"
         End
      End
   End
   Begin VB.Menu mnusep1 
      Caption         =   "Legacy reports"
      Begin VB.Menu mnuReportsCashSales 
         Caption         =   "Front desk sales (by date)"
      End
      Begin VB.Menu mnuFDS 
         Caption         =   "Front desk sales (by batch code)"
      End
      Begin VB.Menu mnuComm 
         Caption         =   "Commissions"
      End
      Begin VB.Menu mnuCGRN 
         Caption         =   "&Consolidated GRN"
      End
   End
   Begin VB.Menu mnusep2 
      Caption         =   "Trading"
      Begin VB.Menu mnuTransfers 
         Caption         =   "&Transfers"
      End
      Begin VB.Menu mnuINVCN 
         Caption         =   "Invoices and Credit notes (by date)"
      End
      Begin VB.Menu mnuSalesItems 
         Caption         =   "&Customer trading (cube)"
      End
      Begin VB.Menu mnuSupplierTrading 
         Caption         =   "&Supplier trading (cube)"
      End
      Begin VB.Menu mnuPOETA 
         Caption         =   "Supplier trading by invoice date/ETA"
      End
      Begin VB.Menu mnuDebtorsReport 
         Caption         =   "Debtors' report"
      End
      Begin VB.Menu mnuTradingPerformance 
         Caption         =   "Trading performance"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTradingPerformanceReport 
         Caption         =   "Trading performance"
      End
      Begin VB.Menu mnuGLAccounts 
         Caption         =   "GL accounts (inventory)"
      End
   End
   Begin VB.Menu mnuMan 
      Caption         =   "Audit reports"
      Begin VB.Menu mnuAdjustmentCube 
         Caption         =   "Adjustment cube"
      End
      Begin VB.Menu mnuUnDELLS 
         Caption         =   "&Unordered deliveries"
      End
      Begin VB.Menu mnuDupEAN 
         Caption         =   "&Duplicate ISBN codes"
      End
      Begin VB.Menu mnuDupSHortCodes 
         Caption         =   "D&uplicate short codes (# codes)"
      End
      Begin VB.Menu mnuNegQty 
         Caption         =   "Qty on hand < 0"
      End
      Begin VB.Menu mnuMissingPrices 
         Caption         =   "Missing prices"
      End
      Begin VB.Menu mnuNonStandardVAT 
         Caption         =   "Non-standard VAT rates"
      End
      Begin VB.Menu mnuReceiving 
         Caption         =   "Receiving rate"
      End
      Begin VB.Menu mnuAuditPoints 
         Caption         =   "Price changes"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTillAudit 
         Caption         =   "Till audit"
      End
      Begin VB.Menu mnuTillJournal 
         Caption         =   "Till journal"
      End
      Begin VB.Menu mnuTillLostScans 
         Caption         =   "Till lost scans"
      End
   End
   Begin VB.Menu mnuStktkeRep 
      Caption         =   "Stocktaking reports"
      Begin VB.Menu mnuStockRecalc 
         Caption         =   "Run stock recalculation (as per dayend)"
      End
      Begin VB.Menu mnuScans 
         Caption         =   "Browse used scan files"
      End
      Begin VB.Menu mnuLastStockTakeList 
         Caption         =   "Last stock-take list"
      End
      Begin VB.Menu mnuDiscrepancyReports 
         Caption         =   "Discrepancy reports"
         Begin VB.Menu mnuDiscrepancyAll 
            Caption         =   "Most recent discrepancy report (all adjustments)"
         End
         Begin VB.Menu mnuDiscrepancyNegOnly 
            Caption         =   "Most recent discrepancy report (-ive adjustments only)"
         End
         Begin VB.Menu mnuDiscrepancyPosOnly 
            Caption         =   "Most recent discrepancy report (+ive adjustments only)"
         End
      End
      Begin VB.Menu mnuStockMove 
         Caption         =   "Stock movements"
      End
      Begin VB.Menu mnuReportsSupplierList 
         Caption         =   "S&upplier List"
      End
      Begin VB.Menu mnuReportsReorderStock 
         Caption         =   "Reorder &Stock"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPOS 
      Caption         =   "Sales reports"
      Begin VB.Menu mnuACRec 
         Caption         =   "Account payments received"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTopSales 
         Caption         =   "&Top sales"
      End
      Begin VB.Menu mnuSalesDet 
         Caption         =   "Sales details"
      End
      Begin VB.Menu mnuSalesbyWeek 
         Caption         =   "Sales by week"
      End
      Begin VB.Menu mnuSalesByMonth 
         Caption         =   "Sales by month"
      End
      Begin VB.Menu mnuSalesSummary1 
         Caption         =   "Sales details by supplier"
      End
      Begin VB.Menu mnuSalesSummary2 
         Caption         =   "Sales details by customer"
      End
      Begin VB.Menu mnuSalesSummary3 
         Caption         =   "Sales details by product"
      End
      Begin VB.Menu mnuSalesByPT 
         Caption         =   "Sales details by product type"
      End
      Begin VB.Menu mnuSalesBySection 
         Caption         =   "Sales details by category"
      End
      Begin VB.Menu mnuSalesByCategory 
         Caption         =   "Sales by category"
      End
      Begin VB.Menu mnuInvoiceSales 
         Caption         =   "Sales by supplier (Inv. & CN only)"
      End
      Begin VB.Menu mnussr 
         Caption         =   "Sales summary report"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSalesCustSumm 
         Caption         =   "Sales by customer summary (Inv. & CN. only)"
         Begin VB.Menu mnuSalesByCustSummDB 
            Caption         =   "Sales by customer (summary DB only)"
         End
         Begin VB.Menu mnuSalesByCustSummCR 
            Caption         =   "Sales by customer (summaryCR only)"
         End
         Begin VB.Menu mnuSalesByCustSummDBCR 
            Caption         =   "Sales by customer (summary DB and CR)"
         End
      End
      Begin VB.Menu mnuSalesByCustBudget 
         Caption         =   "Sales by customer (against budget)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCASHUP 
         Caption         =   "P.O.S. cash-up by period"
      End
      Begin VB.Menu mnuSalesPerformance 
         Caption         =   "Sales performance"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSalesItemsobs 
         Caption         =   "&Customer sales (drill-down)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "&User-designed reports"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuNewUDR 
      Caption         =   "User designed reports"
   End
   Begin VB.Menu mnuExport 
      Caption         =   "&Exports"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oRPT As z_reports

Private nRet         As Long
Private nMainhWnd    As Long

Dim ofrm As Form

Private Type RECT
    Left            As Long
    top             As Long
    Right           As Long
    Bottom          As Long
End Type

' Used to get width and height dimensions for a bitmap
Private Type BITMAP
    bmType          As Long
    bmWidth         As Long
    bmHeight        As Long
    bmWidthBytes    As Long
    bmPlanes        As Integer
    bmBitsPixel     As Integer
    bmBits          As Long
End Type

'Used to get the dimensions of the MDIClient area
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'We need to use this to get the MDIClient area's device context to draw on (and to release it later)
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
'Used to manipulate the GDI32 objects we create / use
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Used to create either a solid or texture brush, and then fill the rectangular area
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Used for drawing the logo in the middle of our MDIClient area
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Used to get the system color, just in case the user turned the background texture off
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long


'''''''''''''''''''''''''''''

'Private mclsMDI      As New clsMDIBackground
Private mlngPrevIndex As Long
Private Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" (ByVal hWnd&, _
    ByVal lpClassName$, ByVal nMaxCount&) As Long
    
Dim strStaffName As String
Dim dteFrom As Date
Dim dteTo As Date
Dim rs As ADODB.Recordset
Dim bForceClose As Boolean

Public Property Let ForceClose(bForce As Boolean)
    bForceClose = bForce
End Property

Private Sub MDIForm_Load()
    On Error GoTo errHandler
Dim strError As String

    GetThunder
    PaintFirstScreen
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub PaintFirstScreen()
Dim fs As New FileSystemObject
Dim oTF As New z_TextFile
    Me.BackColor = RGB(36, 60, 140)
    If UBound(arCommandLine()) > 0 Then
        If arCommandLine(1) <> "N" Then
            BackColor = vbRed
        End If
    Else
    End If
    If oPC.DatabaseName = "PBKS_TEST" Then
        BackColor = vbRed
    End If
    mnuFDS.Visible = Not (oPC.GetProperty("DeliveryStyle") = "BB")
    mnuReportsCashSales.Visible = Not (oPC.GetProperty("DeliveryStyle") = "BB")
    Caption = "Papyrus II Reports v1.3"
    Me.SB1.Panels(1) = "Last day-end: " & oPC.Configuration.LastUpdateDateF & "   "
    Me.SB1.Panels(2) = "   " & oPC.NewQuotation
    Me.SB1.Panels(3) = "   " & IIf(oPC.DatabaseName <> "PBKS", "Server:" & oPC.servername & " Database:" & oPC.DatabaseName, "Server:" & oPC.servername)
    SB1.Panels(2).ToolTipText = SB1.Panels(2).text
    If Not fRunningInIde Then
        subclassMDIClientArea Me
        DrawLogo GetProp(Me.hWnd, "MAINhMDIClient")
    End If
    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\DUMMY1.XLS") Then
        oTF.OpenTextFileToAppend oPC.SharedFolderRoot & "\Templates\DUMMY1.XLS"
        oTF.CloseTextFile
    End If

End Sub
Private Sub GetThunder()
    On Error GoTo errHandler
Dim hIcon As Long
    
    nRet = GetWindowLong(Me.hWnd, GWL_HWNDPARENT)
    Do While nRet
       nMainhWnd = nRet
       nRet = GetWindowLong(nMainhWnd, GWL_HWNDPARENT)
    Loop
    ' set the icon
    Set Me.Icon = Picture1.Picture
    ' get a handle to ICON_BIG
    hIcon = SendMessage(Me.hWnd, WM_GETICON, ICON_BIG, ByVal 0)
    ' send ICON_BIG to the main window
    SendMessage nMainhWnd, WM_SETICON, ICON_BIG, ByVal hIcon

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetThunder"
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    
    If Not bForceClose Then
        If UnloadMode = 0 Or UnloadMode = 1 Then
            If MsgBox("You want to close Papyrus II Reports?", vbQuestion + vbYesNo, "Application closing") = vbNo Then
                Cancel = True
            End If
    End If
        LogSaveToFile "Form count = " & Forms.Count
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode), EA_NORERAISE
    HandleError
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo errHandler
    For Each ofrm In Forms
        Unload ofrm
    Next
    Set ofrm = Nothing
'    Set mclsMDI = Nothing
    Set frmMain = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCustom_Click()
    On Error GoTo errHandler
Dim frm As New frmCustom
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCustom_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuAdjustmentCube_Click()
Dim frm As New frmAdjustmentsCube
    frm.Show

End Sub

Private Sub mnuAgedAppros_Click()
    On Error GoTo errHandler

    Set ofrm = New frmAgedAppros
    ofrm.Component "Aged appros outstanding", True
    ofrm.Show


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAgedAppros_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuApprosIssued_Click()
    On Error GoTo errHandler
    Set ofrm = New frmAppros
    ofrm.Component "ALL"
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReportsApprosIssued_Click", , EA_NORERAISE
    HandleError
End Sub


'Private Sub mnuAuditPoints_Click()
'Dim rpt As DDActiveReports2.ActiveReport
'Dim tf As New z_TextFile
'Dim s As String
'Dim oSQL As New z_SQL
'Dim oMD As New z_ReportMetadata
'Dim sTableName As String
'Dim sec As DDActiveReports2.Section
'Dim ctl As DDActiveReports2.DataControl
'
'    Set rpt = New DDActiveReports2.ActiveReport
'    tf.OpenTextFileToRead oPC.SharedFolderRoot & "\Ad-hoc reports specifications\AuditPoints.XML"
'    s = tf.ReadWholeFile
'    tf.CloseTextFile
'    oMD.LoadMetadataToXML s
'
'    oMD.ConnectionString = oPC.ConnectionString
'    'oMD.RecordSource = "(SELECT *,ISNULL(L_Description, '') + ' - ' + ISNULL(L_MainAuthor, '') as DescriptionCombo FROM " & sTableName & ")"
'
'    rpt.LoadLayout StringToByteArray(oMD.Layout_fromXML, False, True)
'   ' rpt.Show
'    rpt.PrintReport True
'    Set rpt = Nothing
'
'End Sub

Private Sub mnuCASHUP_Click()
Dim frm As New frmPOSSummary

    'If blind cashup is in use, then find oldest unsigned cashup and report until (not including) then with message
    frm.Show
    
End Sub

Private Sub mnuCGRN_Click()
Dim frm As New frmCGRN
    frm.Show
End Sub



Private Sub mnuCOIN_Click()
Dim frm As New frmCOI

    frm.Show
    

End Sub
Private Sub mnuAged_Click()
Dim frm As New frmAgedStock
    
    frm.Show
    
End Sub


Private Sub mnuCOIST_Click()
Dim lngSTID As Long
Dim rs As New ADODB.Recordset
Dim strDate As String
Dim bExVat As Boolean

Dim frm As New frmReportRepresentation

    frm.Component True
    frm.Show vbModal

    Set rs = oPC.CO.execute("SELECT MAX(STKTKE_ID) FROM tSTKTKE JOIN tTR ON STKTKE_ID = TR_ID WHERE TR_STATUS IN (3,4)")
    If rs.State <> 0 Then
        If Not rs.EOF Then
            lngSTID = CLng(rs.Fields(0))
            rs.Close
            Set rs = Nothing
        End If
    End If
    If lngSTID > 0 Then
        Set rs = oPC.CO.execute("SELECT STKTKE_CUTOFFDATE FROM tSTKTKE JOIN tTR ON STKTKE_ID = TR_ID WHERE TR_ID = " & lngSTID)
        If rs.State <> 0 Then
            If Not rs.EOF Then
                strDate = Format(rs.Fields(0), "dd/m/yyyy Hh:Nn")
                rs.Close
                Set rs = Nothing
            End If
        End If
    End If
    If lngSTID > 0 Then
'        If MsgBox("Do you want to show this report with values Ex VAT?", vbQuestion + vbYesNo, "Ex V.A.T. status") = vbYes Then
'            bExVat = True
'        Else
'            bExVat = False
'        End If
        Screen.MousePointer = vbHourglass
        Set oRPT = New z_reports
    'oRPT.COI bExVAT, frm.ReportPresentation
        oRPT.COI_Stocktake_Adj lngSTID, strDate, frm.ExVAT, frm.ReportPresentation, (frm.chkLDP = 1)
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnuComm_Click()
Dim frm As New frmCommissionDialogue
Dim zr As New z_reports

    If SecurityControl(enSECURITY_COMM_AUTH, , "Entering commissions", "You do not have permission to open the commission report.") = False Then Exit Sub
    frm.Show vbModal
    If Not frm.CancelReport Then
        zr.Commissions frm.STAFFID, frm.dtpFrom, frm.dtpTo, frm.ReportPresentation
    End If
   Unload frm
    
End Sub

Private Sub mnuDailyValues_Click()
Dim frmbetween As New frmBetweendates
    frmbetween.Component DateAdd("m", -1, Date), Date
    frmbetween.Visible = True
    frmbetween.Show
End Sub


Private Sub mnuDebtorsReport_Click()
Dim oRPT As New z_reports
Dim f As New frmReportRepresentation
Dim bExVat As Boolean
Dim enPresentation As enumReportPresentation
Dim bRecordsReturned As Boolean
Dim rs As New ADODB.Recordset
Dim rpt As arDebtors

    f.Component False, False
    f.Show vbModal
    enPresentation = f.ReportPresentation
    bExVat = f.ExVAT
    Unload f
    
    Screen.MousePointer = vbHourglass
    oRPT.DebtorsReport rs, bRecordsReturned
    Set rpt = New arDebtors
    rpt.Component rs
    oRPT.PresentReport rpt, f.ReportPresentation, "Debtors", 1000, 200, 12000
    Screen.MousePointer = vbDefault





End Sub

Private Sub mnuDiscrepancyNegOnly_Click()
Dim oRPT As New z_reports
Dim f As New frmReportRepresentation
Dim bExVat As Boolean
Dim enPresentation As enumReportPresentation

    f.Show vbModal
    enPresentation = f.ReportPresentation
    bExVat = f.ExVAT
    Unload f
    
    Screen.MousePointer = vbHourglass
    oRPT.DiscrepancyReport "NEG", bExVat, False, enPresentation
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuDiscrepancyPosOnly_Click()
Dim oRPT As New z_reports
Dim f As New frmReportRepresentation
Dim bExVat As Boolean
Dim enPresentation As enumReportPresentation

    f.Show vbModal
    enPresentation = f.ReportPresentation
    bExVat = f.ExVAT
    Unload f

    Screen.MousePointer = vbHourglass
    oRPT.DiscrepancyReport "POS", bExVat, False, enPresentation
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuDiscrepancyAll_Click()
Dim oRPT As New z_reports
Dim f As New frmReportRepresentation
Dim bExVat As Boolean
Dim enPresentation As enumReportPresentation

    f.Show vbModal
    enPresentation = f.ReportPresentation
    bExVat = f.ExVAT
    Unload f

    Screen.MousePointer = vbHourglass
    oRPT.DiscrepancyReport "ALL", bExVat, False, enPresentation
    Screen.MousePointer = vbDefault

End Sub



Private Sub mnuDupEAN_Click()
Dim oRPT As New z_reports


    Screen.MousePointer = vbHourglass
    oRPT.DuplicateEAN enPreview
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuDupSHortCodes_Click()
Dim oRPT As New z_reports


    Screen.MousePointer = vbHourglass
    oRPT.DuplicateShortCodes enPreview
    Screen.MousePointer = vbDefault
End Sub

'Private Sub mnuDeliveries_Click()
'Dim frm As frmTopSales_dlg
'Dim oRpt As z_reports
'Dim blnNoRecsReturned As Boolean
'Dim strErrMsg As String
'Dim lngTPID As Long
'Dim lngPTID As Long
'Dim dte1 As Date
'Dim dte2 As Date
'Dim strSupplierName As String
'Dim strPTName As String
'
'    On Error GoTo Err_Handler
'    Set frm = New frmTopSales_dlg
'    frm.Caption = "Value of deliveries"
'    frm.Show vbModal
'    dte1 = frm.StartDate
'    dte2 = frm.EndDate
'    lngTPID = frm.SupplierID
'    lngPTID = frm.PTID
'    strSupplierName = frm.SupplierName
'    strPTName = frm.PTName
'    Set oRpt = New z_reports
'
'    strErrMsg = oRpt.ValueOfDeliveries(dte1, dte2, lngTPID, lngPTID, blnNoRecsReturned, strSupplierName, strPTName)
'    Unload frm
'    If strErrMsg > "" Then
'        MsgBox strErrMsg, vbOKOnly, "ERROR"
'    ElseIf blnNoRecsReturned Then
'        MsgBox "No records returned", vbOKOnly, "Papyrus Reports"
'    End If
'EXIT_Handler:
'    Set oRpt = Nothing
'    Exit Sub
'Err_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
'    Resume
'
'End Sub

Private Sub mnuExit_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExit_Click", , EA_NORERAISE
    HandleError
End Sub

Friend Sub DrawLogo(hWnd As Long)
    On Error GoTo errHandler

    Dim aDC        As Long          ' Device context of the MDIClient area
    Dim rcClient   As RECT          ' RECT structure with dimension of MDIClient area
    Dim aPic       As StdPicture    ' Logo picture for center of MDIClient area
    Dim aMask      As StdPicture    ' Mask image so we can draw the logo transparent
    Dim picDC      As Long          ' temporary DC to hold the picture image in
    Dim maskDC     As Long          ' temporary DC to hold the mask image in
    Dim oldBmp1    As Long          ' original 1x1 bitmap for the temporary picDC
    Dim oldBmp2    As Long          ' original 1x1 bitmap for the temporary maskDC
    
    Dim backDC     As Long          ' back buffer device context.
    Dim backBmp    As Long          ' back buffer bitmap
    Dim aBmp       As BITMAP        ' bitmap used to get the picture's dimensions
    Dim abrush     As Long          ' Brush used to paint the background of the MDIClient area
    Dim x          As Long          ' X location for drawing our logo picture
    Dim Y          As Long          ' Y location for drawing our logo picture

    ' Get the MDIClient area's device context
    aDC = GetDC(hWnd)
    ' Get the MDIClient dimensions
    GetWindowRect hWnd, rcClient
    ' shift the origin to 0,0
    rcClient.Right = rcClient.Right - rcClient.Left
    rcClient.Bottom = rcClient.Bottom - rcClient.top
    rcClient.top = 0
    rcClient.Left = 0

    ' Create a backbuffer so we can draw in memory first, then transfer the
    '  background to the MDIClient area all at once.
    backDC = CreateCompatibleDC(aDC)
    backBmp = CreateCompatibleBitmap(aDC, rcClient.Right, rcClient.Bottom)
    DeleteObject SelectObject(backDC, backBmp)

    'Paint window background
'    If chkBGTexture.Value = 0 Then
        ' Use the system setting for application workspace
        If UBound(arCommandLine) > 0 Then
            If arCommandLine(1) <> "N" Then
                abrush = CreateSolidBrush(vbRed)
            Else
                abrush = CreateSolidBrush(RGB(25, 38, 85))
            End If
        Else
                abrush = CreateSolidBrush(RGB(25, 38, 85))
        End If
           'Me.BackColor = RGB(36, 60, 140)

 '   Else
        ' Create a pattern brush using the background texture
 '       abrush = CreatePatternBrush(imgBG.Picture.Handle)
 '   End If
    ' Fill the backbuffer with the selected brush
    FillRect backDC, rcClient, abrush
    ' Clean up our brush object
    DeleteObject abrush

    ' Do logo, if that has been selected.
'    If chkLogo.Value = 1 Then
        Set aPic = imgLogo.Picture
        Set aMask = imgLogoMask.Picture
        ' Get logo's dimensions - overkill? Probably, but I HATE screwing around
        '  with himetric units. They make me want to kick something really really
        '  hard. And you wouldn't want me to break my toe, would you? :-p
        GetObject aPic.Handle, Len(aBmp), aBmp
        ' Create some compatible device contexts to hold our logo pics in
        picDC = CreateCompatibleDC(aDC)
        maskDC = CreateCompatibleDC(aDC)
        ' Select our pictures into the temporary DCs, and keep a reference to
        '  the original 1x1 bitmaps so we can replace them later, freeing our logo images.
        oldBmp1 = SelectObject(picDC, aPic.Handle)
        oldBmp2 = SelectObject(maskDC, aMask.Handle)
        ' Calculate the x and y location for our logo
        x = (rcClient.Right - aBmp.bmWidth) ' \ 2
        Y = (rcClient.Bottom - aBmp.bmHeight) ' \ 2
        ' punch the hole for our logo
        BitBlt backDC, x, Y, aBmp.bmWidth, aBmp.bmHeight, maskDC, 0, 0, vbMergePaint
        ' draw the logo
        BitBlt backDC, x, Y, aBmp.bmWidth, aBmp.bmHeight, picDC, 0, 0, vbSrcAnd
        
        ' Replace the original 1x1 bitmaps (which frees our logo pictures)
        SelectObject picDC, oldBmp1
        SelectObject maskDC, oldBmp2
        ' Clean up the graphics objects
        DeleteDC picDC
        DeleteObject oldBmp1
        DeleteDC maskDC
        DeleteObject oldBmp2
 '   End If
    
    ' blt from backbuffer into client rectangle - Transfers the entire thing at once.
    BitBlt aDC, 0, 0, rcClient.Right, rcClient.Bottom, backDC, 0, 0, vbSrcCopy
    ' Clean up our backbuffer objects
    DeleteDC backDC
    DeleteObject backBmp
    ' Release our hold on the device context
    ReleaseDC hWnd, aDC

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DrawLogo(hwnd)", hWnd
End Sub
Private Function fRunningInIde() As Boolean
    On Error GoTo errHandler
Dim sClassName As String
Dim nStrLen    As Long

    '
    ' See if we're running in the IDE.
    '
    sClassName = String$(260, vbNullChar)
    nStrLen = GetClassName(Me.hWnd, sClassName, Len(sClassName))
    If nStrLen Then sClassName = Left$(sClassName, nStrLen)
    
    fRunningInIde = (sClassName = "ThunderMDIForm")
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.fRunningInIde"
End Function

Private Sub mnuPrintCustomerOrders_Click()
    On Error GoTo errHandler
    Set ofrm = New frmCO
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPrintCustomerOrders_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuExport_Click()
'Dim frm As frmExports
'    Set frm = New frmExports
'    frm.Show vbModal
'
'    Unload frm
End Sub



Private Sub mnuGLAccounts_Click()
'Dim f As frmGLAccounts
'
'    f.Show
    
End Sub

Private Sub mnuINVCN_Click()
Dim oRPT As New z_reports
Dim frm As New frmPeriodDialogue

    frm.Show vbModal
    
    If frm.CancelReport Then Exit Sub
    Screen.MousePointer = vbHourglass
    oRPT.Invoices_CreditNotes frm.FromDate, frm.ToDate, frm.ReportPresentation
    Screen.MousePointer = vbDefault
    
    Unload frm

End Sub


Private Sub mnuLastStockTakeList_Click()
Dim oRPT As New z_reports
Dim frm As New frmLastStockTakeList

    frm.Show
End Sub

Private Sub mnuMissingPrices_Click()
Dim frm As New frmMissingPricesDialog

    frm.Show
    
End Sub

Private Sub mnuNegQty_Click()
Dim frm As New frmTrackingNegativeStock

    frm.Show
    
End Sub

Private Sub mnuNewUDR_Click()
Dim frmUDR As New frmUDR

    frmUDR.Show
    
End Sub

Private Sub mnuNonStandardVAT_Click()
    On Error GoTo errHandler
Dim oRPT As New z_reports
Dim rs As ADODB.Recordset
Dim rpt As arNonStandardVAT
Dim frm As New frmReportRepresentation

    frm.Component False, False
    frm.Show vbModal
    
    If frm.Cancelled Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    oRPT.NonStandardVAT rs
    Set rpt = New arNonStandardVAT
    rpt.Component rs
    oRPT.PresentReport rpt, frm.ReportPresentation, "Non standard VAT ", 1000, 200, 12000
    Unload frm
    Screen.MousePointer = vbDefault


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNonStandardVAT_Click"
End Sub


Private Sub mnuReceiving_Click()
Dim frm As New frmPeriodDialogue
'Dim frmPT As New frmReceivingPT
Dim dte1 As Date
Dim dte2 As Date
Dim strSQL As String
Dim rs As ADODB.Recordset


'    frm.Show vbModal
'
'    If frm.CancelReport Then Exit Sub
'    Screen.MousePointer = vbHourglass
'
'    dte1 = frm.FromDate
'    dte2 = frm.ToDate
'    Unload frm
'
'    strSQL = "SELECT * FROM vReceivingRate WHERE dte BETWEEN '" & ReverseDate(dte1) & "' AND '" & ReverseDate(dte2) & "'"
'    Set rs = New ADODB.Recordset
'
'    Me.SB1.Panels(1).text = "Loading . . . "
'    DoEvents
'    rs.Open strSQL, oPC.CO
'    Screen.MousePointer = vbDefault
'    Set frmPT = New frmReceivingPT
'    frmPT.Component rs, "Receiving rate"
'    frmPT.Show 'vbModal
'    Me.SB1.Panels(1).text = ""
'    Set rs = Nothing


End Sub

Private Sub mnuReportsApprosOS_Click()
    On Error GoTo errHandler
    Set ofrm = New frmAppros
    ofrm.Component "OS"
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReportsApprosOS_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuReportsCashSales_Click()
    On Error GoTo errHandler
    Set ofrm = New frmCashSales
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReportsCashSales_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuFDS_Click()
    Set ofrm = New frmFrontDeskSales
    ofrm.Show
    Exit Sub
End Sub
Private Sub mnuInvoiceSales_Click()
    Set ofrm = New frmInvoiceSales
    
    ofrm.Show

End Sub

Private Sub mnuReportsDeliveries_Click()
    On Error GoTo errHandler
    Set ofrm = New frmDeliveries
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReportsDeliveries_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuReportsHashNumbers_Click()
    On Error GoTo errHandler
Dim oRPT As z_reports
Dim blnNoRecsReturned As Boolean
Dim strErrMsg As String
    If MsgBox("Print list of all hash numbers?", vbYesNo + vbQuestion, "Papyrus II Reports") = vbNo Then
        Exit Sub
    End If
    
    Set oRPT = New z_reports
    strErrMsg = oRPT.HashNumbers(blnNoRecsReturned)
    If strErrMsg > "" Then
        MsgBox strErrMsg, vbOKOnly, "ERROR"
    ElseIf blnNoRecsReturned Then
        MsgBox "No records returned", vbOKOnly, "Papyrus Reports"
    End If
EXIT_Handler:
    Set oRPT = Nothing
'Err_Handler:
'    MsgBox Error
'    GoTo EXIT_Handler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReportsHashNumbers_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuRerun_Click()
Dim f As New frmSelectStocktake
    f.Show vbModal
    
End Sub

Private Sub mnuSalesByCategory_Click()
    Dim frm As New frmSalesByPeriod
    
    frm.Component "Sales by category", "CAT"
    frm.Show

End Sub

Private Sub mnuSalesByCustSummDB_Click()
Dim frm As New frmDebtors_all
    frm.Component "INVOICE SALES"
    frm.Show
    
 End Sub
Private Sub mnuSalesByCustSummDBCR_Click()
Dim frm As New frmDebtors_all
    frm.Component ""
    frm.Show
    
 End Sub
Private Sub mnuSalesByCustSummCR_Click()
Dim frm As New frmDebtors_all
    frm.Component "CREDIT RETURNS"
    frm.Show
 End Sub
Private Sub mnuSalesByCustBudget_Click()
'Dim oRPT As z_reports
'Dim frm As New frmPeriodDialogue
'Dim rs As ADODB.Recordset
'Dim rpt As arSalesByCustomerBudget
'
'    frm.Show vbModal
'    If frm.CancelReport Then
'        Unload frm
'        DoEvents
'        Exit Sub
'    End If
'    dteFrom = frm.dtpFrom
'    dteTo = frm.dtpTo
'
'    Screen.MousePointer = vbHourglass
'    Set oRPT = New z_reports
'    oRPT.SalesByCustomerBudget rs, frm.FromDate, frm.ToDate
'    Set rpt = New arSalesByCustomerBudget
'
'    rpt.Component rs, dteFrom, dteTo
'    oRPT.PresentReport rpt, frm.ReportPresentation, "SalesByCustomer", 1000, 200, 12000
'
'    Unload frm
'    Screen.MousePointer = vbDefault
'
 End Sub

Private Sub mnuSalesByPT_Click()
    On Error GoTo errHandler
    
Dim frm As New frmSalesByPeriod

    frm.Component "Sales by product type", "PT"
    frm.Show
    
'Dim oRPT As z_reports
'Dim frm As New frmPeriodDialogue
'Dim dteFrom As Date
'Dim dteTo As Date
'Dim rpt As arSalesByPTByDate
'Dim rs As ADODB.Recordset
'
'    frm.Show vbModal
'    If frm.CancelReport Then
'        Unload frm
'        DoEvents
'        Exit Sub
'    End If
'    dteFrom = frm.dtpFrom
'    dteTo = frm.dtpTo
'    Unload frm
'
'    Screen.MousePointer = vbHourglass
'
'    Set oRPT = New z_reports
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseClient
'    oRPT.SalesByPTByDate rs, dteFrom, dteTo
'    Set rpt = New arSalesByPTByDate
'
'    rpt.Component rs, "Sales by product type from " & Format(dteFrom, "dd/mm/yyyy") & " to " & Format(dteTo, "dd/mm/yyyy"), Format(dteTo, "dd/mm/yyyy"), "From " & Format(dteFrom, "dd/mm/yyyy") & " to " & Format(dteTo, "dd/mm/yyyy")
'    oRPT.PresentReport rpt, frm.ReportPresentation, "SalesByCustByDate", 1000, 200, 12000
'
'    Set oRPT = Nothing
'
'    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSalesByPT_Click"
End Sub

Private Sub mnuSalesBySection_Click()
    Dim frm As New frmSalesByPeriod
    
    frm.Component "Sales by category", "CAT"
    frm.Show
    
End Sub

Private Sub mnuSalesDet_Click()
Dim frm As New frmSalesDetails

    frm.Show

End Sub

Private Sub mnuSalesbyWeek_Click()
Dim frm As New frmSalesByPeriod

    frm.Component "Sales by week", "W"
    frm.Show
End Sub
Private Sub mnuSalesbyMonth_Click()
Dim frm As New frmSalesByPeriod
    
    frm.Component "Sales by month", "M"
    frm.Show

End Sub


Private Sub mnuSalesSummary1_Click()
Dim oRPT As z_reports
Dim frm As New frmSalesBySupplier

    frm.Show
    
End Sub
Private Sub mnuSalesSummary2_Click()
Dim frm As New frmSalesByCustomer

    frm.Show
    
End Sub
Private Sub mnuSalesItems_Click()
    On Error GoTo errHandler
Dim lngTPID As Long
Dim frm As frmCustomersTA
Dim strSQL As String
Dim rs As ADODB.Recordset
Dim frmR As frmCustomerPT
Dim dte1 As Date
Dim dte2 As Date
Dim oSQL As z_SQL
Dim dteLimitToView As Date

    Set frm = New frmCustomersTA
    frm.Component "Customer trading filters"
    frm.Show vbModal
    If frm.Cancelled Then
        Unload frm
        DoEvents
        Exit Sub
    End If
    dte1 = frm.StartDate
    dte2 = frm.EndDate
    
    
    If oPC.BlindCashup = True Then
        Set oSQL = New z_SQL
        dteLimitToView = oSQL.GetDateOfEarliestUnSignedSession
        If EndOfDay(dte2) > dteLimitToView Then
            MsgBox "There are unsigned cash ups starting prior to your selected end date (" & Format(dteLimitToView, "dd/mm/yyyy") & "). You cannot include thse in the report. Select an earlier end date.", vbInformation, "Can't do this"
            Exit Sub
        End If
    End If
'    If oPC.BlindCashup Then
'        Set oSQL = New z_SQL
'        dteLastAllowableSessionDate = oSQL.GetDateOfMostRecentFullySignedOffDailySession(dte1)
'        If dte2 > dteLastAllowableSessionDate Then
'            MsgBox "There are unsigned cash ups prior to your selected end date. These cannot be included in the report. Select an earlier end date.", vbInformation, "Can't do this"
'            Exit Sub
'        End If
'    End If
'
    
    
    lngTPID = frm.CustomerID
    If lngTPID > 0 Then
        strSQL = "SELECT * FROM zCustomerPT_UNION WHERE   TPID = " & lngTPID & " AND dte BETWEEN dbo.StartofDay('" & ReverseDate(dte1) & "') AND dbo.endofDay('" & ReverseDate(dte2) & "')"
    Else
        strSQL = "SELECT * FROM zCustomerPT_UNION WHERE dte BETWEEN dbo.StartofDay('" & ReverseDate(dte1) & "') AND dbo.endofDay('" & ReverseDate(dte2) & "')"
    End If
    Set rs = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Me.SB1.Panels(1).text = "Loading . . . "
    DoEvents
    rs.Open strSQL, oPC.CO
    Screen.MousePointer = vbDefault
    Set frmR = New frmCustomerPT
    frmR.Component rs, "Customer", Not frm.chkExVAT = 1, IIf(frm.chkLDP = 1, "LPD", "")
    Me.SB1.Panels(1).text = ""
    Unload frm
    frmR.Show 'vbModal
    Set rs = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSalesItems_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSTA_Click()
Dim oRPT As New z_reports
    Screen.MousePointer = vbHourglass
    oRPT.StockTakeAdjustments enPreview
    Screen.MousePointer = vbDefault
    
End Sub





Private Sub mnuSalesSummary3_Click()
Dim frm As New frmSalesSummaryPerProduct

    frm.Component "Sales summary by product", 2, DateAdd("m", -1, Date), CDate(0), "B", True

    frm.Show

End Sub

Private Sub mnuScans_Click()
Dim frm As New frmBrowseScanFiles

    frm.Show
    
End Sub


Private Sub mnuStockList_Click()
Dim frm As New frmStockDialogue
    frm.Show vbModal
    
End Sub

Private Sub mnuStockMove_Click()
Dim oRPT As New z_reports
Dim frm As New frmPeriodDialogue

    frm.Show vbModal
    
    If frm.CancelReport Then Exit Sub
    Screen.MousePointer = vbHourglass
    oRPT.StockMovements frm.FromDate, frm.ToDate, frm.ReportPresentation
    Screen.MousePointer = vbDefault
    
    Unload frm

End Sub

Private Sub mnuStockRecalc_Click()
10        On Error GoTo errHandler
      Dim SQL As New z_SQL
    Screen.MousePointer = vbHourglass
20        SQL.RunProc "RecalcSOH", Array(), "running recalcSOH"
    Screen.MousePointer = vbDefault
30        Exit Sub
errHandler:
40        If ErrMustStop Then Debug.Assert False: Resume
50        ErrorIn "frmMain.mnuStockRecalc_Click"
End Sub

Private Sub mnuStockRecon_Click()
'Dim F As New frmReconciliation
'    If oPC.Configuration.SignTransactions = True Then
'        If SecurityControl(enSECURITY_ISSUPERVISOR, , "You must be a supervisor to view this. Please sign", DOCACCESS, , strStaffName) = False Then
'               Exit Sub
'        Else
'            F.Component strStaffName
'            F.Show
'        End If
'    Else
'        F.Show
'    End If
End Sub

Private Sub mnuSupplierTrading_Click()
Dim frm As frmSuppliersCube

    Set frm = New frmSuppliersCube
    frm.Show
    
EXIT_Handler:
    Exit Sub
errHandler:
    ErrPreserve
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSupplierTrading_Click", , EA_NORERAISE
    HandleError

End Sub

Private Sub mnuPOETA_Click()
Dim frm As frmSuppliersCube

    Set frm = New frmSuppliersCube
    frm.Component True
    frm.Show

End Sub

Private Sub mnuSwapDB_Click()
Dim f As Form

    Screen.MousePointer = vbHourglass
    For Each f In Forms
        If Not f Is Forms(0) Then Unload f
    Next
    
    oPC.SwapConnectionToDatabase
    PaintFirstScreen
    If oPC.DatabaseName = "PBKS_TEST" Then
        Me.mnuSwapDB.Caption = "Swap to working on LIVE database"
    Else
        Me.mnuSwapDB.Caption = "Swap to working on TEST database"
    End If
    Screen.MousePointer = vbDefault
    
    MsgBox "You are now connected to the " & IIf(oPC.DatabaseName = "PBKS_TEST", "TEST", "LIVE") & " database", vbOKOnly, "Status"

End Sub

Private Sub mnuTillAudit_Click()
Dim frm As New frmTillAudit
    frm.Show

End Sub

Private Sub mnuTillJournal_Click()
Dim frm As New frmTillJournal
    frm.Show
End Sub

Private Sub mnuTillLostScans_Click()
Dim frm As New frmTillLostScans
    frm.Show

End Sub

Private Sub mnuTradingPerformance_Click()
Dim f As New frmTradingPerformance

    f.Show
    
    
End Sub


Private Sub mnuTradingPerformanceReport_Click()
Dim f As New frmTradingPerformanceReport

    f.Show

End Sub

Private Sub mnuTransfers_Click()

Dim frm As frmProductPT

    Set frm = New frmProductPT
    frm.Show

EXIT_Handler:
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTransfers_Click", , EA_NORERAISE
    HandleError

End Sub

Private Sub mnuTopSales_Click()
Dim frm As frmTopSaless

    Set frm = New frmTopSaless
    frm.Show
    

End Sub

Private Sub mnuReportsInvoices_Click()
    On Error GoTo errHandler
    Set ofrm = New frmInvoices
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReportsInvoices_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuReportsReorderStock_Click()
    On Error GoTo errHandler
    Set ofrm = New frmReorderStock
    ofrm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReportsReorderStock_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub mnuReportsSeeSafe_Click()
'    Set ofrm = New frmSeeSafe
'    ofrm.Show
'End Sub

Private Sub mnuReportsSupplierList_Click()
    On Error GoTo errHandler
Dim oRPT As z_reports
Dim blnNoRecsReturned As Boolean
Dim strErrMsg As String

    If MsgBox("Print list of all suppliers on record?", vbYesNo + vbQuestion, "Papyrus Reports") = vbNo Then
        Exit Sub
    End If
    
    Set oRPT = New z_reports
    strErrMsg = oRPT.TradingPartners(2, blnNoRecsReturned)  '   2 = Supplier role type
    If strErrMsg > "" Then
        MsgBox strErrMsg, vbOKOnly, "ERROR"
    ElseIf blnNoRecsReturned Then
        MsgBox "No records returned", vbOKOnly, "Papyrus Reports"
    End If
    Set oRPT = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReportsSupplierList_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuUnDELLS_Click()
Dim oRPT As New z_reports
Dim frm As New frmPeriodDialogue

    frm.Show vbModal
    
    If frm.CancelReport Then Exit Sub
    Screen.MousePointer = vbHourglass
    oRPT.UnorderedDeliveryLines frm.FromDate, frm.ToDate, frm.ReportPresentation
    Screen.MousePointer = vbDefault
    
    Unload frm

End Sub

'
'Private Sub mnuUser_Click()
'    On Error GoTo errHandler
'Dim frm As New frmUserDesign
'    frm.Show
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuUser_Click", , EA_NORERAISE
'    HandleError
'End Sub
'
