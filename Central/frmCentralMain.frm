VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00552619&
   Caption         =   "Papyrus Central"
   ClientHeight    =   6825
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "frmCentralMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar PB1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   6390
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13124
            MinWidth        =   13124
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   6059
            MinWidth        =   6068
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      Picture         =   "frmCentralMain.frx":058A
      ScaleHeight     =   0
      ScaleWidth      =   9600
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9600
      Begin VB.Image imgLogo 
         Height          =   1800
         Left            =   1200
         Picture         =   "frmCentralMain.frx":0B14
         Top             =   -75
         Width           =   4755
      End
      Begin VB.Image imgLogoMask 
         Height          =   1800
         Left            =   7305
         Picture         =   "frmCentralMain.frx":1C996
         Top             =   120
         Width           =   4755
      End
   End
   Begin MSComctlLib.Toolbar TBHEAD 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   345
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList3"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   23
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bpo"
            Object.ToolTipText     =   "Browse purchase orders"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bco"
            Object.ToolTipText     =   "Browse customer orders"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "binv"
            Object.ToolTipText     =   "Browse invoices"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bcn"
            Object.ToolTipText     =   "Browse credit notes"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bdel"
            Object.ToolTipText     =   "Browse deliveries"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bapp"
            Object.ToolTipText     =   "Browse appros"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bappr"
            Object.ToolTipText     =   "Browse appro returns"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btr"
            Object.ToolTipText     =   "Browse transfers"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bcs"
            Object.ToolTipText     =   "Browse cash sales"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "npo"
            Object.ToolTipText     =   "New purchase order"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nco"
            Object.ToolTipText     =   "New customer order"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ninv"
            Object.ToolTipText     =   "New invoice"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ncn"
            Object.ToolTipText     =   "New credit note"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ndel"
            Object.ToolTipText     =   "New delivery"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "napp"
            Object.ToolTipText     =   "New appro"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nappr"
            Object.ToolTipText     =   "New appro return"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ntr"
            Object.ToolTipText     =   "New transfer"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bbks"
            Object.ToolTipText     =   "Browse books"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "bGen"
            Object.ToolTipText     =   "Browse general stock"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nbk"
            Object.ToolTipText     =   "New book"
            ImageIndex      =   25
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ngs"
            Object.ToolTipText     =   "New general stock"
            ImageIndex      =   19
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3315
      Top             =   5070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3881A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":38DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3934E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":398E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":39E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3A41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3A9B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3AF50
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3B4EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3BA84
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3C01E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3C5B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3CB52
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3D0EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3D686
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3DC20
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3E1BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3E754
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3ECEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3F288
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3F822
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":3FDBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":40356
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":408F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentralMain.frx":40D42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuSaveColumnWidths 
         Caption         =   "Save column widths"
      End
      Begin VB.Menu mnuBu 
         Caption         =   "Backup"
         Begin VB.Menu mnuBULocal 
            Caption         =   "Backup to local drive"
         End
         Begin VB.Menu mnuBURemovable 
            Caption         =   "Backup to removable"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuBrowse 
      Caption         =   "&Browse"
      Begin VB.Menu mnuBrowseProducts 
         Caption         =   "P&roducts"
      End
      Begin VB.Menu mnuCashups 
         Caption         =   "Cash-ups"
      End
      Begin VB.Menu mnuCustomerS 
         Caption         =   "&Customers"
      End
      Begin VB.Menu mnuExchanges 
         Caption         =   "&Exchanges"
      End
      Begin VB.Menu mnuCOLS 
         Caption         =   "Sales orders outstanding"
      End
      Begin VB.Menu mnuUploads 
         Caption         =   "&Uploads"
      End
      Begin VB.Menu mnDownloads 
         Caption         =   "&Downloads"
      End
      Begin VB.Menu mnuPromotions 
         Caption         =   "&Promotions"
      End
      Begin VB.Menu mnuBrowseSuppliers 
         Caption         =   "&Suppliers"
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "&New"
      Begin VB.Menu mnuNewCustomer 
         Caption         =   "&Customer"
      End
      Begin VB.Menu mnuProduct 
         Caption         =   "&Product"
      End
      Begin VB.Menu mnuSupplier 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlert 
         Caption         =   "Customer alert"
      End
   End
   Begin VB.Menu mnuClassifications 
      Caption         =   "&Classifications"
      Begin VB.Menu mnuDictionary 
         Caption         =   "&Dictionary"
      End
      Begin VB.Menu mnuPT 
         Caption         =   "&Product types"
      End
   End
   Begin VB.Menu mnuMaster 
      Caption         =   "&Master files"
      Begin VB.Menu mnuConfig 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuCountries 
         Caption         =   "Countri&es"
      End
      Begin VB.Menu mnuStores 
         Caption         =   "&Stores"
      End
      Begin VB.Menu mnuBudgets 
         Caption         =   "&Budgets"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "R&eports"
      Begin VB.Menu mnuSalesP 
         Caption         =   "Sales patterns"
      End
      Begin VB.Menu mnuSalesPerf 
         Caption         =   "Sales performance"
      End
      Begin VB.Menu mnuPrepSalesData1 
         Caption         =   "Prepare data for sales spreadsheets"
      End
      Begin VB.Menu mn 
         Caption         =   "Print edited customer list"
      End
      Begin VB.Menu mnuInvalidAcno 
         Caption         =   "Exchanges with invalid Ac/No's"
      End
      Begin VB.Menu mnuReminders 
         Caption         =   "&Reminders"
      End
      Begin VB.Menu mnuSTAT1 
         Caption         =   "&Statistics"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Actions"
      Begin VB.Menu mnuWash 
         Caption         =   "Wash against Nielsen"
      End
      Begin VB.Menu mnuApprove 
         Caption         =   "Approve changes to customer database"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import"
         Begin VB.Menu mnuBulkDeliveryImport 
            Caption         =   "&Bulk delivery import"
         End
         Begin VB.Menu mnuImportLoyalty 
            Caption         =   "&Customers and sales"
         End
         Begin VB.Menu mnuBIC 
            Caption         =   "&BIC codes from Bookfind"
         End
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Begin VB.Menu mnuExpSBLoyalty 
            Caption         =   "Export Loyalty editing results to branches"
         End
         Begin VB.Menu mnuExportLC 
            Caption         =   "Export edited customer records"
         End
         Begin VB.Menu mnuExportAll 
            Caption         =   "Export all customer records"
         End
         Begin VB.Menu mnuEXCat 
            Caption         =   "Items on &catalogue(s)"
         End
         Begin VB.Menu mnuWants 
            Caption         =   "&Wants"
         End
         Begin VB.Menu mnuDBScript 
            Caption         =   "&Database script"
         End
      End
      Begin VB.Menu mnuBrStatusRep 
         Caption         =   "Branch synchronization"
         Begin VB.Menu mnuCustomerStatusReport 
            Caption         =   "Loyalty customer status report"
         End
         Begin VB.Menu mnuLoyaltyCustomerMatch 
            Caption         =   "Loyalty customer record audit"
         End
         Begin VB.Menu mnuSalesMatchAudit 
            Caption         =   "Exchanges audit"
         End
         Begin VB.Menu mnuSOHBulk 
            Caption         =   "Request stock on hand bulk update"
         End
         Begin VB.Menu mnuFetchCashupSets 
            Caption         =   "Request cashup sets"
         End
         Begin VB.Menu mnuRQCOLS 
            Caption         =   "Request sales orders from branches"
         End
      End
      Begin VB.Menu mnuUtilities 
         Caption         =   "&Utilities"
         Begin VB.Menu mnuDmpSUPP 
            Caption         =   "Search &suppliers"
         End
         Begin VB.Menu mnuDmpCUST 
            Caption         =   "Search &customers"
         End
         Begin VB.Menu mnuDmpPROD 
            Caption         =   "Search &products"
         End
         Begin VB.Menu mnuMerge 
            Caption         =   "&Merge two products"
         End
         Begin VB.Menu mnuMergeCust 
            Caption         =   "M&erge two customers"
         End
         Begin VB.Menu mnuMergePT 
            Caption         =   "Merge two& product types"
         End
      End
      Begin VB.Menu mnuDiag 
         Caption         =   "&Diagnostics"
      End
      Begin VB.Menu mnuSendCustChanges 
         Caption         =   "Send customer changes to branches"
      End
      Begin VB.Menu mnuSB 
         Caption         =   "Service broker"
      End
   End
   Begin VB.Menu mnuBranchData 
      Caption         =   "Branch data"
      Begin VB.Menu mnuBRCountries 
         Caption         =   "Countries"
      End
   End
   Begin VB.Menu mnuMailing 
      Caption         =   "&Mailing"
      Begin VB.Menu mnuCustMail 
         Caption         =   "&Manage customer mailings"
         Begin VB.Menu mnuCustomerBrowseContext 
            Caption         =   "CustomerBrowseMenu"
            Begin VB.Menu mnuAlertHistory 
               Caption         =   "Alert history"
            End
         End
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuBrowseCustomerPopup 
      Caption         =   "Lists"
      Visible         =   0   'False
      Begin VB.Menu mnuAddtoList 
         Caption         =   "Add to current list"
      End
      Begin VB.Menu mnuRemoveFromList 
         Caption         =   "Remove from list"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nRet         As Long
Private nMainhWnd    As Long

Private Type RECT
    Left            As Long
    Top             As Long
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
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
'We need to use this to get the MDIClient area's device context to draw on (and to release it later)
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
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

Public frmMainCustomerPreview As frmCustomerPreview
Public frmMainLoyaltyPreview As frmLoyaltyPreview
Public WithEvents oBF As zc_BF
Attribute oBF.VB_VarHelpID = -1

'''''''''''''''''''''''''''''
Dim frmBrowseProd As frmBrowseProducts
Private mlngPrevIndex As Long
Private Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" (ByVal hwnd&, _
    ByVal lpClassName$, ByVal nMaxCount&) As Long

Dim frmBrowseCustomers As frmBrowseCustomersEx
Enum EnumMode
    enEditingRow = 0
    enAddingRow = 1
    enNotEditing = 3
End Enum

Private Sub mnuAlert_Click()
Dim f As New frmAlert
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim strCustname As String
Dim strCustAcno As String
Dim lngTPID As Long

    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    strCustname = frm.CustomerName
    strCustAcno = frm.Accnum
    
    Unload frm
    If lngTPID = 0 Then Exit Sub
    f.Component lngTPID, strCustname, strCustAcno
    f.Show
End Sub

Private Sub mnuAlertHistory_Click()
    Me.ActiveForm.mnuAlertHistory
End Sub


Private Sub mnuBrowseSuppliers_Click()
Dim frm As New frmBrowsesuppliers
    
    frm.Show
    
End Sub

Private Sub mnuBudgets_Click()
Dim f As New frmBudgetManagement
    f.Show
    
End Sub

Private Sub mnuBulkDeliveryImport_Click()
Dim frm As New frmDeliveryImport

    frm.Show
    
    
End Sub

Private Sub mnuBULocal_Click()
    On Error GoTo errHandler
Dim oBU As New z_Backup
Dim ret As Boolean

    Screen.MousePointer = vbHourglass
    ret = oBU.BackupToLocal
    Screen.MousePointer = vbDefault
    MsgBox "Backup done", , "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBULocal_Click"
    HandleError
End Sub
Private Sub mnuBURemovable_Click()
    On Error GoTo errHandler
Dim oBU As New z_Backup
Dim ret As Boolean

    Screen.MousePointer = vbHourglass
    ret = oBU.BackupToLocal
    If ret = False Then MsgBox "Problem backing up database", , "Warning"
    ret = oBU.ZIPBackupToNonLocal
    If ret = False Then MsgBox "Problem zipping backup file", , "Warning"
    Screen.MousePointer = vbDefault
    MsgBox "Backup done", , "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBURemovable_Click"
    HandleError
End Sub


Private Sub mnuCustomerStatUpdateRequest_Click()
Dim f As New frmStoreSelection
    f.Show
End Sub

Private Sub mnuCashups_Click()
Dim f As New frmBrowseCashups
    
    f.Show
    
End Sub

Private Sub mnuCOLS_Click()
Dim f As New frmBrowseCOLS
    
    f.Show
    

End Sub

Private Sub mnuCustomerStatusReport_Click()
Dim f As New frmBranchStatsReport

    f.Show vbModal
    
End Sub

Private Sub mnuFetchCashupSets_Click()
Dim f As New frmStoreSelectionForCashupResend

    f.Component "Request cashup data", "Cashup"
    f.Show
    
End Sub

Private Sub mnuLoyaltyCustomerMatch_Click()
Dim f As New frmBranchMatchReport
    
    f.Show
    
End Sub

Private Sub mnuPrepSalesData1_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
Dim f As New frmPeriodDialogue
Dim frm As New frmProductPT
Dim rs As ADODB.Recordset
Dim z As New z_StockManager
Dim dteFrom As Date
Dim dteTo As Date
    f.Component "Select sales period", 2, Date, False
    f.Show vbModal
    dteFrom = f.dtpFrom
    dteTo = f.dtpTo
    
    If f.CancelReport = True Then
        Unload f
        Exit Sub
    End If
    
    Unload f
    
    Screen.MousePointer = vbHourglass
    oSQL.PrepareSalesSpreadsheetData dteFrom, dteTo
    Screen.MousePointer = vbDefault
    MsgBox "Data is prepared for spreadsheets. Open Excel and open the reporting spreadsheets and refresh the data.", , "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPrepSalesData1_Click"
    HandleError
End Sub

Private Sub mnuRQCOLS_Click()
Dim f As New frmStoreSelectionForCashupResend

    f.Component "Request sales order data", "COLS"
    f.Show

End Sub

Private Sub mnuSalesMatchAudit_Click()
Dim f As New frmBranchMatchSalesReport
    
    f.Show

End Sub

Private Sub mnuSalesPerf_Click()
    On Error GoTo errHandler
Dim f As New frmPeriodDialogue1
Dim frm As New frmSalesPT
Dim rs As ADODB.Recordset
Dim z As New z_StockManager
Dim dteFrom As Date
Dim dteTo As Date

    f.Component "Select  period", 2, Date, False
    f.Show vbModal
    dteFrom = f.dtpFrom
    dteTo = f.dtpTo
    
    If f.CancelReport = True Then
        Unload f
        Exit Sub
    End If
    
    Unload f
    
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    z.GetSalesPerformance dteFrom, dteTo, rs
    frm.Component rs
    Screen.MousePointer = vbDefault
    frm.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSalesP_Click"
    HandleError

End Sub

Private Sub mnuSOHBulk_Click()
Dim f As New frmStoreSelection

    f.Component "Stock on hand request", "SOHB"
    f.Show vbModal
    MsgBox "Wait a while (could be a minute or more) and then click 'Refresh'.", vbInformation, "Status"
    Unload f

End Sub

Private Sub mnuStores_Click()
Dim f As New frmStoreManagement
    f.Show

End Sub

Private Sub mnuSupplier_Click()
    NewSupplier
End Sub
Private Sub NewSupplier()
    On Error GoTo errHandler
Dim frm As frmSupplier
Dim oSupp As a_Supplier
    Set frm = New frmSupplier
    Set oSupp = New a_Supplier
    frm.Component oSupp
    frm.Show
    Exit Sub
errHandler:
    ErrorIn "frmMain.NewSupplier"
End Sub

Private Sub oBF_Progress(lngPos As Long, lngMax As Long)
    On Error GoTo errHandler
    If lngPos Mod 100 = 0 Then
        Me.SB1.Panels(2).Text = "       Record " & CStr(lngPos) & " of " & CStr(lngMax)
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oBF_Progress(lngPos,lngMax)", Array(lngPos, lngMax)
    HandleError
End Sub
Private Sub oBF_ProgressB(lngPos As Long, lngMax As Long, pMsg As String)
    On Error GoTo errHandler
    'If lngPos Mod 100 = 0 Then
        Me.SB1.Panels(2).Text = pMsg & CStr(lngPos) & " of " & CStr(lngMax)
    'End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oBF_ProgressB(lngPos,lngMax,pMsg)", Array(lngPos, lngMax, pMsg)
    HandleError
End Sub

Private Sub oBF_Status(msg As String)
    On Error GoTo errHandler
    If msg = "DUP" Then
        MsgBox "The operation cannot complete because there are duplicate product codes in the database." _
        & vbCrLf & "Use the Reports application to report the duplicates, correct them, then restart this operation.", vbInformation, "Can't do this"
    Else
        Me.SB1.Panels(2).Text = msg
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oBF_Status(msg)", msg
    HandleError
End Sub

Private Sub Dictionary()
    On Error GoTo errHandler
Dim frm As frmDictionary
    Set frm = New frmDictionary
    frm.Show 'vbModal
    Set frm = Nothing

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Dictionary"
End Sub
Private Sub NewBook()
    On Error GoTo errHandler
Dim frmA As frmProduct
Dim frm As frmProduct
Dim oProd As a_Product

    Set oProd = Constructor.CreateProduct(False)

    Set frm = New frmProduct
    frm.Component oProd
    frm.Show
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewBook"
End Sub
Private Sub NewGenStock()
    On Error GoTo errHandler
Dim frm As frmProductNB
Dim oProd As a_Product

    Set oProd = Constructor.CreateProduct(True)

    Set frm = New frmProductNB
    frm.Component oProd
    frm.Show
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewGenStock"
End Sub

Private Sub NewCustomer()
    On Error GoTo errHandler
Dim frm As frmCustomer
Dim oCust As a_Customer
    Set oCust = New a_Customer
    oCust.BeginEdit
    oCust.InitializeNewCustomer
    Set frm = New frmCustomer
    frm.Component oCust
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewCustomer"
End Sub
'Private Sub NewSupplier()
'Dim frm As frmSupplier
'Dim oSupp As a_Supplier
'    Set frm = New frmSupplier
'    Set oSupp = New a_Supplier
'    frm.Component oSupp
'    frm.Show
'End Sub
'Private Sub BrowseInvoices()
'    If frmBrowseInvoices Is Nothing Then
'       Set frmBrowseInvoices = New frmBrowseInvoices
'    End If
'    frmBrowseInvoices.ZOrder 0
'End Sub
'Private Sub BrowseReturns()
'    If frmBrowseReturns Is Nothing Then
'       Set frmBrowseReturns = New frmBrowseReturns
'    End If
'    frmBrowseReturns.ZOrder 0
'End Sub
'
'Private Sub BrowseDELS()
'    If frmBrowseDEL Is Nothing Then
'       Set frmBrowseDEL = New frmBrowseDels
'    End If
'    frmBrowseDEL.ZOrder 0
'End Sub
'Private Sub BrowseTrans()
'    If frmBrowseTF Is Nothing Then
'       Set frmBrowseTF = New frmBrowseTF
'    End If
'    frmBrowseTF.ZOrder 0
'End Sub
'Private Sub BrowsePOs()
'    If frmBrowsePO Is Nothing Then
'       Set frmBrowsePO = New frmBrowsePOs
'    End If
'    frmBrowsePO.ZOrder 0
'
'End Sub
'Private Sub BrowseOrders()
'    If frmBrowseCO Is Nothing Then
'       Set frmBrowseCO = New frmBrowseCOs
'    End If
'    frmBrowseCO.ZOrder 0
'End Sub
'Private Sub BrowseCS()
'    If frmBrowseCS Is Nothing Then
'       Set frmBrowseCS = New frmBrowseCS
'    End If
'    frmBrowseCS.ZOrder 0
'End Sub
'Private Sub BrowseCN()
'    If frmBrowseCN Is Nothing Then
'       Set frmBrowseCN = New frmBrowseCN
'    End If
'    frmBrowseCN.ZOrder 0
'End Sub
'Private Sub BrowseApps()
'    If frmBrowseAPP Is Nothing Then
'       Set frmBrowseAPP = New frmBrowseAPPs
'    End If
'    frmBrowseAPP.ZOrder 0
'End Sub
'Private Sub BrowseAPPRs()
'    If frmBrowseAPPR Is Nothing Then
'       Set frmBrowseAPPR = New frmBrowseAPPRs
'    End If
'    frmBrowseAPPR.ZOrder 0
'End Sub
Private Sub BrowseBooks()
    On Error GoTo errHandler
    If frmBrowseProd Is Nothing Then
        Set frmBrowseProd = New frmBrowseProducts
    End If
    frmBrowseProd.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseBooks"
End Sub
Private Sub BrowseGenStock()
    On Error GoTo errHandler
    If frmBrowseGS Is Nothing Then
        Set frmBrowseGS = New frmBrowseGS
    End If
    frmBrowseGS.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseGenStock"
End Sub

Friend Sub BrowseCustomers()
    On Error GoTo errHandler
    If frmBrowseCustomers Is Nothing Then
       Set frmBrowseCustomers = New frmBrowseCustomersEx
    End If
    frmBrowseCustomers.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseCustomers"
End Sub


Private Sub MDIForm_Load()
    On Error GoTo errHandler
Dim strError As String
    GetThunder
    If Command() <> "" Then
        BackColor = vbRed
    Else
        Me.BackColor = RGB(36, 60, 140)
    End If
    If oPC.BFLoaded Then
        Me.Caption = "Papyrus Central    -    Connected to Bookfind"
    Else
        Me.Caption = "Papyrus Central"
    End If
    Me.SB1.Panels(2) = oPC.ServerName
If Not fRunningInIde Then
    subclassMDIClientArea Me
    DrawLogo GetProp(Me.hwnd, "MAINhMDIClient")
End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_Load"
    HandleError
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    If UnloadMode = 0 Or UnloadMode = 1 Then
    If MsgBox("You want to close Wordsworth Central?", vbQuestion + vbYesNo, "Application closing") = vbNo Then
        Cancel = True
    End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode)
    HandleError
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo errHandler
    unsubclassMDIClientArea Me
    Set frmMain = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_Unload(Cancel)", Cancel
    HandleError
End Sub


Private Sub mn_Click()
    On Error GoTo errHandler
Dim frm As frmPrintedited
    Set frm = New frmPrintedited
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mn_Click"
    HandleError
End Sub

Private Sub mnuAddtoList_Click()
    On Error GoTo errHandler
    Me.ActiveForm.AddToList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAddtoList_Click"
    HandleError
End Sub

Private Sub mnuApprove_Click()
    On Error GoTo errHandler
Dim frm As frmTPApprove
    Set frm = New frmTPApprove
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuApprove_Click"
    HandleError
End Sub

Private Sub mnuBrowseProducts_Click()
    On Error GoTo errHandler
        Screen.MousePointer = vbHourglass
        If frmBrowseProd Is Nothing Then
            Set frmBrowseProd = New frmBrowseProducts
        End If
        frmBrowseProd.ZOrder 0
        Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseProducts_Click"
    HandleError
End Sub

Private Sub mnuExchanges_Click()
    On Error GoTo errHandler
Dim frm As New frmExchanges
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExchanges_Click"
    HandleError
End Sub

Private Sub mnuExportAll_Click()
    On Error GoTo errHandler
Dim oEX As z_Import
Dim iRecordsExported As Long
    If MsgBox("You are exporting ALL the customers.", vbOKCancel + vbQuestion, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set oEX = New z_Import
    iRecordsExported = oEX.AppendEditedCustomers("LCE_ALL" & Format(Now, "yyyymmddHHNN") & ".TXT", CDate(0), CDate(0), True)
    Set oEX = Nothing
    Screen.MousePointer = vbDefault
    MsgBox "Export of customers complete." & vbCrLf & "Records exported : " & CStr(iRecordsExported), , "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExportAll_Click"
    HandleError
End Sub

Private Sub mnuExportLC_Click()
    On Error GoTo errHandler
Dim frm As New frmExportLC
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExportLC_Click"
    HandleError
End Sub



Private Sub mnuInvalidAcno_Click()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
Dim f As New frmSpuriousAcnos
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------

    rs.CursorLocation = adUseClient
    
    rs.Open "SELECT dbo.tEXCHANGE.CUSTOMERACNO, dbo.tEXCHANGE.EXCHANGENUMBER, dbo.tEXCHANGE.EXCHANGEDATE, dbo.tEXCHANGE.BRANCHCODE " _
          & " FROM dbo.tEXCHANGE LEFT OUTER JOIN dbo.tTP ON dbo.tEXCHANGE.CUSTOMERACNO = dbo.tTP.TP_ACNo WHERE     (dbo.tTP.TP_ACNo IS NULL) ORDER BY CUSTOMERACNO,EXCHANGENUMBER", oPC.COShort, adOpenKeyset
    f.Component rs
    f.Show vbModal
    
    Unload f
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuInvalidAcno_Click"
    HandleError
End Sub

Private Sub mnuRemoveFromList_Click()
    On Error GoTo errHandler
    Me.ActiveForm.RemoveFromList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRemoveFromList_Click"
    HandleError
End Sub

Private Sub mnuBIC_Click()
    On Error GoTo errHandler
Dim frm As New frmBICImport
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBIC_Click"
    HandleError
End Sub

Private Sub mnuBrowseCustomers_Click()
    On Error GoTo errHandler
    BrowseCustomers
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseCustomers_Click"
    HandleError
End Sub


Private Sub mnuBrowseStock_Click()
    On Error GoTo errHandler
    BrowseGenStock
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseStock_Click"
    HandleError
End Sub










Private Sub mnuConfig_Click()
    On Error GoTo errHandler
Dim frm As frmConfiguration
'    If SecurityControl(4, gSTAFFID, , "Changing the configuration requires security level 4.") = False Then Exit Sub
    Set frm = New frmConfiguration
    frm.Component oPC.Configuration
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuConfig_Click"
    HandleError
End Sub

Private Sub mnuCountries_Click()
    On Error GoTo errHandler
Dim frm As frmCountry
    Set frm = New frmCountry
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCountries_Click"
    HandleError
End Sub


Private Sub mnuCustMail_Click()
    On Error GoTo errHandler
Dim frmMail As New frmMailing
    frmMail.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCustMail_Click"
    HandleError
End Sub

Private Sub mnuCustomers_Click()
    On Error GoTo errHandler
    If frmBrowseCustomers Is Nothing Then
       Set frmBrowseCustomers = New frmBrowseCustomersEx
    End If
    frmBrowseCustomers.ZOrder 0

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCustomers_Click"
    HandleError
End Sub

Private Sub mnuDictionary_Click()
    On Error GoTo errHandler
'    If SecurityControl(3, gSTAFFID, , "Enter your security code.", "You do not have permission to edit the dictionary.") = False Then Exit Sub

    Dictionary
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDictionary_Click"
    HandleError
End Sub

'Private Sub mnuDmpCUST_Click()
'Dim frm As New frmExportCUST
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
''    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
''        MsgBox "You do not have security to export records.", vbExclamation, "Denied"
''        Exit Sub
''    End If
'    frm.Show vbModal
'End Sub
'
'Private Sub mnuDmpPROD_Click()
'Dim frm As New frmExportPROD
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
''    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
''        MsgBox "You do not have security to export records.", vbExclamation, "Denied"
''        Exit Sub
''    End If
'    frm.Show vbModal
'End Sub
'
'Private Sub mnuDmpSUPP_Click()
'Dim frm As New frmExportSUPP
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
''    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
''        MsgBox "You do not have security to export records.", vbExclamation, "Denied"
''        Exit Sub
''    End If
'    frm.Show vbModal
'End Sub

Private Sub mnuExit_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExit_Click"
    HandleError
End Sub



Private Sub mnuImportLoyalty_Click()
    On Error GoTo errHandler
Dim oImp As New z_Import
Dim f, fc, fol
Dim oFSO As New FileSystemObject
Dim strMsg As String
Dim strLog As String

    If MsgBox("You want to import loyalty customers from the FTP site?", vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    oImp.ConnectToFTP
    
    Me.SB1.Panels(1).Text = "Fetching files from FTP site . . ."
    oImp.FetchLoyaltyFiles strMsg
    
    oImp.ConfirmReceipt
    
    oImp.CloseFTP
    oImp.Hangup
    
    Me.SB1.Panels(1).Text = "Updating database from imported files  . . ."
    oImp.UpdateDBFromFiles strLog
    
    'Delete files from \PBKS\Data\Loyalty\up
    Set fol = oFSO.GetFolder(oPC.SharedFolderRoot & "\Data\Loyalty\UP")
    Set fc = fol.files
    For Each f In fc
        f.Delete
    Next

    Screen.MousePointer = vbDefault
    Me.SB1.Panels(1).Text = ""
    MsgBox "File import completed and updated." & vbCrLf & "results as follows: " & vbCrLf & vbCrLf & strMsg & vbCrLf & vbCrLf & strLog, , "Status"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportLoyalty_Click"
    HandleError
End Sub
Private Sub mnuExpSBLoyalty_Click()
    On Error GoTo errHandler
Dim oImp As New z_Import
    If MsgBox("You want to export loyalty customers editing results to the branches?", vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    
    oImp.SendLoyaltyEditingChangesToBranches
    

    Screen.MousePointer = vbDefault
    MsgBox "Export procedure launched.", vbInformation, "Status"
    Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExpSBLoyalty_Click"
    HandleError
End Sub
Private Sub mnuMerge_Click()
    On Error GoTo errHandler
Dim frm As frmMergeProducts
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
'        MsgBox "You do not have security to merge products.", vbExclamation, "Denied"
'        Exit Sub
'    End If
    Set frm = New frmMergeProducts
    frm.Show vbModal
    Unload frm
    Set frm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMerge_Click"
    HandleError
End Sub

Private Sub mnuMergeCust_Click()
    On Error GoTo errHandler
Dim frm As frmMergeTPs
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
'        MsgBox "You do not have security to merge customers or suppliers.", vbExclamation, "Denied"
'        Exit Sub
'    End If
    Set frm = New frmMergeTPs
    frm.Show vbModal
    Unload frm
    Set frm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMergeCust_Click"
    HandleError
End Sub

Private Sub mnuMergePT_Click()
    On Error GoTo errHandler
'Dim frm As New frmMergePTs
'    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMergePT_Click"
    HandleError
End Sub

Private Sub mnuNBP_Click()
    On Error GoTo errHandler
    NewGenStock
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNBP_Click"
    HandleError
End Sub

Private Sub mnuNewCustomer_Click()
    On Error GoTo errHandler
    NewCustomer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewCustomer_Click"
    HandleError
End Sub


Private Sub mnuNewStock_Click()
    On Error GoTo errHandler
    NewBook
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewStock_Click"
    HandleError
End Sub

'Private Sub mnuNNS_Click()
'    NewNonStock
'End Sub
'
'Private Sub mnuNonStock_Click()
'    NonStock
'End Sub


Private Sub mnuPT_Click()
    On Error GoTo errHandler
Dim frm As frmPTs
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
'        MsgBox "You do not have security to edit product types.", vbExclamation, "Denied"
'        Exit Sub
'    End If

    Set frm = New frmPTs
    frm.Show 'vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPT_Click"
    HandleError
End Sub




Private Sub mnuSalesP_Click()
    On Error GoTo errHandler
Dim f As New frmPeriodDialogue
Dim frm As New frmProductPT
Dim rs As ADODB.Recordset
Dim z As New z_StockManager
Dim dteFrom As Date
Dim dteTo As Date

    f.Component "Select sales period", 2, Date, False
    f.Show vbModal
    dteFrom = f.dtpFrom
    dteTo = f.dtpTo
    
    If f.CancelReport = True Then
        Unload f
        Exit Sub
    End If
    
    Unload f
    
    Screen.MousePointer = vbHourglass
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    z.GetSalesPatterns dteFrom, dteTo, rs
    frm.Component rs
    Screen.MousePointer = vbDefault
    frm.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSalesP_Click"
    HandleError
End Sub
Private Sub mnuSaveColumnWidths_Click()
    On Error GoTo errHandler

    Me.ActiveForm.mnuSaveLayout
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSaveColumnWidths_Click"
    HandleError
End Sub

'''Private Sub TBHead_ButtonClick(ByVal Button As MSComctlLib.Button)
'''    Select Case UCase(Button.Key)
''''    Case "BINV"
''''        BrowseInvoices
''''    Case "BAPP"
''''        BrowseApps
''''    Case "BAPPR"
''''        BrowseAPPRs
''''    Case "BDEL"
''''        BrowseDELS
''''    Case "BPO"
''''        BrowsePOs
''''    Case "BCO"
''''        BrowseOrders
''''    Case "BCN"
''''        BrowseCN
''''    Case "BGEN"
''''        BrowseGenStock
'''    Case "BBKS"
'''        BrowseBooks
''''    Case "BTR"
''''        BrowseTrans
''''    Case "BCUST"
''''        BrowseCustomers
''''    Case "BSUPP"
''''        BrowseSupplier
''''    Case "BCS"
''''        BrowseCS
''''    Case "NINV"
''''        NewInvoice
''''    Case "NCUST"
''''        NewCustomer
''''    Case "NGS"
''''        NewGenStock
''''    Case "NBK"
''''        NewBook
''''    Case "NSUPP"
''''        NewSupplier
''''    Case "NAPP"
''''        NewAPP
''''    Case "NAPPR"
''''        NewAPPR
''''    Case "NDEL"
''''        NewDEL
''''    Case "NPO"
''''        NewPO
''''    Case "NCO"
''''        NewCO
''''    Case "NCN"
''''        NewCN
''''    Case "NTR"
''''        NewTRANS
''''    Case "NCN"
''''        NewCN
'''    End Select
'''End Sub
Friend Sub DrawLogo(hwnd As Long)
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
    aDC = GetDC(hwnd)
    ' Get the MDIClient dimensions
    GetWindowRect hwnd, rcClient
    ' shift the origin to 0,0
    rcClient.Right = rcClient.Right - rcClient.Left
    rcClient.Bottom = rcClient.Bottom - rcClient.Top
    rcClient.Top = 0
    rcClient.Left = 0

    ' Create a backbuffer so we can draw in memory first, then transfer the
    '  background to the MDIClient area all at once.
    backDC = CreateCompatibleDC(aDC)
    backBmp = CreateCompatibleBitmap(aDC, rcClient.Right, rcClient.Bottom)
    DeleteObject SelectObject(backDC, backBmp)

    'Paint window background
'    If chkBGTexture.Value = 0 Then
        ' Use the system setting for application workspace
           'Me.BackColor = RGB(36, 60, 140)
        If Command() <> "" Then
            abrush = CreateSolidBrush(vbRed)
        Else
            abrush = CreateSolidBrush(RGB(25, 38, 85))
        End If

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
        x = (rcClient.Right - aBmp.bmWidth - 25) ' \ 2
        Y = (rcClient.Bottom - aBmp.bmHeight - 20) ' \ 2
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
    ReleaseDC hwnd, aDC
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.DrawLogo(hwnd)", hwnd
End Sub
Private Function fRunningInIde() As Boolean
    On Error GoTo errHandler
Dim sClassName As String
Dim nStrLen    As Long

    '
    ' See if we're running in the IDE.
    '
    sClassName = String$(260, vbNullChar)
    nStrLen = GetClassName(Me.hwnd, sClassName, Len(sClassName))
    If nStrLen Then sClassName = Left$(sClassName, nStrLen)
    
    fRunningInIde = (sClassName = "ThunderMDIForm")
  
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.fRunningInIde"
End Function

'''''''''''''''''''''''''
Private Sub GetThunder()
    On Error GoTo errHandler
Dim hIcon As Long
    
    nRet = GetWindowLong(Me.hwnd, GWL_HWNDPARENT)
    Do While nRet
       nMainhWnd = nRet
       nRet = GetWindowLong(nMainhWnd, GWL_HWNDPARENT)
    Loop
    ' set the icon
    Set Me.Icon = Picture1.Picture
    ' get a handle to ICON_BIG
    hIcon = SendMessage(Me.hwnd, WM_GETICON, ICON_BIG, ByVal 0)
    ' send ICON_BIG to the main window
    SendMessage nMainhWnd, WM_SETICON, ICON_BIG, ByVal hIcon

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.GetThunder"
End Sub

Private Sub mnuSB_Click()
    On Error GoTo errHandler
Dim frm As New frmTransmissionControl
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSB_Click"
    HandleError
End Sub

Private Sub mnuSendCustChanges_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.SendCustomerChanges
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSendCustChanges_Click"
    HandleError
End Sub

Private Sub mnuWash_Click()
    On Error GoTo errHandler
Dim frm As frmWash

    If Not InStr(1, oPC.Configuration.LookupSeq, "BF") > 0 Then
        MsgBox "This application does not use Nielsen"
        Exit Sub
    End If
    If MsgBox("This procedure will take some hours and should be run when trading has ended." & vbCrLf & "Do you wish to continue?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
        Set frm = New frmWash
        frm.Show vbModal
        If frm.Cancelled Then
            Unload frm
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        PB1.Visible = True
        Set oBF = New zc_BF
Dim OpenResult As Integer
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
        oBF.UpdateFromBookfind frm.Author = 1, _
        frm.Title = 1, frm.Subtitle = 1, _
        frm.Availability = 1, _
        frm.Binding = 1, frm.Edition = 1, frm.SupplierCode = 1, frm.Publishername = 1, frm.SeriesTitle = 1, _
        frm.PublicationDate = 1, _
        frm.UKPrice = 1, frm.RRP = 1, frm.BIC = 1, frm.BookStatus = 1, gSTAFFID, ""
        Screen.MousePointer = vbDefault
        PB1.Visible = False
        Unload frm
        Set oBF = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuWash_Click"
    HandleError
End Sub

Private Sub TBHEAD_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo errHandler
    Select Case UCase(Button.Key)
'    Case "BINV"
'        BrowseInvoices
'    Case "BAPP"
'        BrowseApps
'    Case "BAPPR"
'        BrowseAPPRs
'    Case "BDEL"
'        BrowseDELS
'    Case "BPO"
'        BrowsePOs
'    Case "BCO"
'        BrowseOrders
'    Case "BCN"
'        BrowseCN
'    Case "BGEN"
'        BrowseGenStock
    Case "BBKS"
        BrowseBooks
'    Case "BTR"
'        BrowseTrans
'    Case "BCUST"
'        BrowseCustomers
'    Case "BSUPP"
'        BrowseSupplier
'    Case "BCS"
'        BrowseCS
'    Case "NINV"
'        NewInvoice
'    Case "NCUST"
'        NewCustomer
'    Case "NGS"
'        NewGenStock
'    Case "NBK"
'        NewBook
'    Case "NSUPP"
'        NewSupplier
'    Case "NAPP"
'        NewAPP
'    Case "NAPPR"
'        NewAPPR
'    Case "NDEL"
'        NewDEL
'    Case "NPO"
'        NewPO
'    Case "NCO"
'        NewCO
'    Case "NCN"
'        NewCN
'    Case "NTR"
'        NewTRANS
'    Case "NCN"
'        NewCN
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.TBHEAD_ButtonClick(Button)", Button
    HandleError
End Sub
