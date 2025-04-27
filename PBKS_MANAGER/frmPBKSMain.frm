VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00552619&
   Caption         =   "Papyrus II - Manager"
   ClientHeight    =   8835
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14355
   Icon            =   "frmPBKSMain.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD1 
      Left            =   480
      Top             =   3420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   8505
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3942
            MinWidth        =   3951
            Key             =   "a"
            Object.Tag             =   "a"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17912
            Key             =   "b"
            Object.Tag             =   "b"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   2884
            MinWidth        =   2893
            Key             =   "c"
            Object.Tag             =   "c"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   2880
      Left            =   0
      Picture         =   "frmPBKSMain.frx":058A
      ScaleHeight     =   2820
      ScaleWidth      =   14295
      TabIndex        =   1
      Top             =   330
      Visible         =   0   'False
      Width           =   14355
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   345
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   960
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Image imgLogo 
         Height          =   1800
         Left            =   1200
         Picture         =   "frmPBKSMain.frx":0B14
         Top             =   -75
         Width           =   4755
      End
      Begin VB.Image imgLogoMask 
         Height          =   1800
         Left            =   6420
         Picture         =   "frmPBKSMain.frx":1C998
         Top             =   195
         Width           =   4755
      End
   End
   Begin MSComctlLib.Toolbar TBHEAD 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
            Object.ToolTipText     =   "Browse Goods received notes"
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
            Object.ToolTipText     =   "New counter invoice"
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
            Picture         =   "frmPBKSMain.frx":3881C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":38DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":39350
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":398EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":39E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3A41E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3A9B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3AF52
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3B4EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3BA86
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3C020
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3C5BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3CB54
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3D0EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3D688
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3DC22
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3E1BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3E756
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3ECF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3F28A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3F824
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":3FDBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":40358
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":408F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPBKSMain.frx":40D44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuManDBCopies 
         Caption         =   "Manage database copies (server only)"
         Begin VB.Menu mnuRestoreTest 
            Caption         =   "Restore test database from selected backup file"
         End
         Begin VB.Menu mnuBackupCur 
            Caption         =   "Backup currently connected database"
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTr 
      Caption         =   "&Actions"
      Begin VB.Menu mnuOutlook 
         Caption         =   "&Send to Outlook folder"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "&Email directly "
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEDI 
         Caption         =   "Transmit EDI file"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVoid 
         Caption         =   "&Void document"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "&Cancel document"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCancelINactive 
         Caption         =   "Ca&ncel and/or fulfil incomplete lines on document"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancelLine 
         Caption         =   "Cancel &document line"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFulfil 
         Caption         =   "&Fulfil document line"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelLine 
         Caption         =   "D&elete document line"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateCreditNote 
         Caption         =   "Create credit note"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSalesComm 
         Caption         =   "&Sales commission"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdjust 
         Caption         =   "&Adjust stock level"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMemo 
         Caption         =   "&Memo"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHeader 
         Caption         =   "&Header"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyLines 
         Caption         =   "Copy document lines to Papyrus clipboard"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPastelines 
         Caption         =   "Paste document lines from Papyrus clipboard to open document"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPastelinestoNEW 
         Caption         =   "Paste document lines from Papyrus clipboard to NEW"
         Begin VB.Menu mnuPastelinestoNEWOrder 
            Caption         =   "Order"
         End
         Begin VB.Menu mnuPastelinestoNewInvoice 
            Caption         =   "Invoice"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPastelinestoNEWAppro 
            Caption         =   "Appro"
         End
         Begin VB.Menu mnuPastelinestoNewCounterInvoice 
            Caption         =   "Counter invoice"
         End
         Begin VB.Menu mnuPastelinestoNEWQuotation 
            Caption         =   "Quotation"
         End
         Begin VB.Menu mnuCopytoNewPO 
            Caption         =   "Purchase order"
         End
         Begin VB.Menu mnuPastelinestoNEWPFInvoice 
            Caption         =   "Pro-forma invoice"
         End
         Begin VB.Menu mnuTransferIn 
            Caption         =   "Transfer IN"
         End
         Begin VB.Menu mnuTransferOut 
            Caption         =   "Transfer OUT"
         End
      End
      Begin VB.Menu mnuRevPapClip 
         Caption         =   "Review Papyrus clipboard"
      End
      Begin VB.Menu mnuCopyDoc 
         Caption         =   "Copy current document to new"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Sep10 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDelact 
         Caption         =   "G.R.N. actions"
         Enabled         =   0   'False
         Begin VB.Menu mnuDiscr 
            Caption         =   "Print discrepancy claim"
         End
         Begin VB.Menu mnuCustomerAlloc 
            Caption         =   "Customer allocations"
         End
         Begin VB.Menu mnuBarcodes 
            Caption         =   "Print barcodes"
         End
      End
      Begin VB.Menu sep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveColumnWidths 
         Caption         =   "&Save column widths"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuBrowse 
      Caption         =   "&Browse"
      Begin VB.Menu mnuSepCustomerStuffBrowse 
         Caption         =   "Customers"
         Begin VB.Menu mnuBrowseCustomers 
            Caption         =   "&Customer master records"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuApps 
            Caption         =   "&Appros"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuAPPRs 
            Caption         =   "Appro &returns"
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuCNotes 
            Caption         =   "&Credit notes"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuBrowseGoodsDelivery 
            Caption         =   "Goods delivery"
         End
         Begin VB.Menu mnuBrowseInvoices 
            Caption         =   "&Invoices"
            Shortcut        =   ^I
         End
         Begin VB.Menu mnuCOS 
            Caption         =   "&Orders"
            Shortcut        =   ^O
         End
         Begin VB.Menu mnuBrowsePF 
            Caption         =   "Proforma invoices"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnuBrowseQuotes 
            Caption         =   "&Quotations"
            Shortcut        =   ^Q
         End
         Begin VB.Menu hc1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCOSR 
            Caption         =   "Customer order status reports"
         End
      End
      Begin VB.Menu mnuSepSupplierStuffBrowse 
         Caption         =   "Supplier"
         Begin VB.Menu mnuBrowseSUpp 
            Caption         =   "&Supplier master records"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuSupplierClaims 
            Caption         =   "Claims"
            Shortcut        =   ^W
         End
         Begin VB.Menu mnuDels 
            Caption         =   "&Goods received"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnuPOS 
            Caption         =   "&Purchase orders"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuReturns 
            Caption         =   "&Returns"
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu h1a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "&Transfers"
      End
      Begin VB.Menu mnuBudget 
         Caption         =   "Budget"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuDashboard 
         Caption         =   "Dashboard"
         Shortcut        =   ^E
      End
      Begin VB.Menu h4bb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCS 
         Caption         =   "Cash sales"
      End
      Begin VB.Menu mnuFD 
         Caption         =   "&Front desk activity"
      End
      Begin VB.Menu h2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowseBooks 
         Caption         =   "&Products"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuBrowseStock 
         Caption         =   "&Non-book stock"
      End
      Begin VB.Menu mnuServiceItem 
         Caption         =   "&Service items (e.g. postage)"
      End
      Begin VB.Menu hx 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTextBites 
         Caption         =   "Text bites"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuBrowseCategoryChecks 
         Caption         =   "Category checks"
      End
   End
   Begin VB.Menu mnuNew 
      Caption         =   "&New"
      Begin VB.Menu mnuNewCustomers 
         Caption         =   "Customers"
         Begin VB.Menu mnuNewCustomer 
            Caption         =   "&Customer master record"
            Shortcut        =   +^{F1}
         End
         Begin VB.Menu mnuAPP 
            Caption         =   "&Appro"
            Shortcut        =   +^{F2}
         End
         Begin VB.Menu mnuAPPR 
            Caption         =   "Appro &return"
            Shortcut        =   +^{F3}
         End
         Begin VB.Menu mnuNewInvoice 
            Caption         =   "&Counter invoice"
            Shortcut        =   +^{F4}
         End
         Begin VB.Menu mnuCNote 
            Caption         =   "&Credit note"
            Shortcut        =   +^{F12}
         End
         Begin VB.Menu mnuGDN 
            Caption         =   "Goods delivered"
            Shortcut        =   +^{F5}
         End
         Begin VB.Menu mnuPreInv 
            Caption         =   "Pre-delivery invoice"
            Shortcut        =   +^{F6}
         End
         Begin VB.Menu mnuProforma 
            Caption         =   "&Pro-forma invoice"
            Shortcut        =   +^{F9}
         End
         Begin VB.Menu mnuCO 
            Caption         =   "&Order"
            Shortcut        =   +^{F7}
         End
         Begin VB.Menu mnuQuote 
            Caption         =   "&Quotation"
            Shortcut        =   +^{F8}
         End
      End
      Begin VB.Menu mnuSuppliers 
         Caption         =   "Suppliers"
         Begin VB.Menu mnuNewSupp 
            Caption         =   "&Supplier master record"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuDel 
            Caption         =   "&Goods received"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnuPO 
            Caption         =   "&Purchase order"
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mnuNewReturn 
            Caption         =   "R&eturn"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu mnuSub 
            Caption         =   "&Subscription order"
            Shortcut        =   ^{F6}
         End
      End
      Begin VB.Menu mnuTranOut 
         Caption         =   "&Transfer OUT"
      End
      Begin VB.Menu mnuTFRIn 
         Caption         =   "Transfer &IN"
      End
      Begin VB.Menu h3 
         Caption         =   "-"
         Index           =   333
      End
      Begin VB.Menu mnuNewBook 
         Caption         =   "&Book product"
      End
      Begin VB.Menu mnuNBP 
         Caption         =   "&Non-book product"
      End
      Begin VB.Menu mnuNewServiceItem 
         Caption         =   "&Service item (e.g. postage)"
      End
      Begin VB.Menu h4b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewCategoryCheck 
         Caption         =   "New category check"
      End
      Begin VB.Menu h4c 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportForTransfer 
         Caption         =   "Import new stock for transfer"
      End
   End
   Begin VB.Menu mnuMast 
      Caption         =   "&Settings"
      Begin VB.Menu mnuDictionary 
         Caption         =   "&Dictionary"
      End
      Begin VB.Menu mnuPT 
         Caption         =   "&Product types"
      End
      Begin VB.Menu mnuCatalogues 
         Caption         =   "&Catalogues"
      End
      Begin VB.Menu mnuCathead 
         Caption         =   "Catalogue headings"
      End
      Begin VB.Menu mnuRR 
         Caption         =   "&Rounding rules"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "&Configuration"
      End
      Begin VB.Menu mnuCountries 
         Caption         =   "Countri&es"
      End
      Begin VB.Menu h21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change password"
      End
   End
   Begin VB.Menu mnuOF 
      Caption         =   "Order fulfilment"
      Begin VB.Menu mnuAlloc 
         Caption         =   "&Order fulfilment"
      End
      Begin VB.Menu mnuOrderFulfilmentAppros 
         Caption         =   "Order fulfilment - with &appros"
      End
      Begin VB.Menu mnuInvoicing 
         Caption         =   "Invoicing"
         Begin VB.Menu mnuInvoicesToDeliver 
            Caption         =   "Pre-delivery invoices to deliver on"
         End
         Begin VB.Menu mnuInvoiceCompleteOrders 
            Caption         =   "Completeness requirement: orders - to invoice"
         End
         Begin VB.Menu mnuCompleteOrdersSnagged 
            Caption         =   "Completeness requirement: orders held up"
         End
      End
   End
   Begin VB.Menu mnuOps 
      Caption         =   "&Tracking && reordering"
      Begin VB.Menu mnuBrowseTracking 
         Caption         =   "Browse tracking actions"
      End
      Begin VB.Menu mnuStatusChange 
         Caption         =   "&Supplier status change"
      End
      Begin VB.Menu mnuODPO 
         Caption         =   "&Track purchase orders"
      End
      Begin VB.Menu mnuCOOD 
         Caption         =   "Track &customer orders"
      End
      Begin VB.Menu mnuTrsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuREORDSales 
         Caption         =   "&Reorder from sales and transfers-out"
      End
      Begin VB.Menu mnuPurch 
         Caption         =   "&Purchase for customer orders"
      End
      Begin VB.Menu mnuTrsep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenReturn 
         Caption         =   "&Generate seesafe return"
         Begin VB.Menu mnuAllSuppliers 
            Caption         =   "All suppliers"
         End
         Begin VB.Menu mnuSelSupplier 
            Caption         =   "Selected supplier"
         End
      End
      Begin VB.Menu mnuTrsep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "R&eports"
         Begin VB.Menu mnuReminders 
            Caption         =   "&Reminders"
         End
         Begin VB.Menu mnuODCOStatus 
            Caption         =   "&Customer orders status"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuImport 
         Caption         =   "&Import"
         Begin VB.Menu mnuImportSales 
            Caption         =   "Sales from portable device (.txt files)"
         End
         Begin VB.Menu mnuImportToClipboard 
            Caption         =   "Data to clipboard"
         End
         Begin VB.Menu mnuNewTransfer 
            Caption         =   "Create new transfer from XML file"
         End
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Begin VB.Menu mnuEXCat 
            Caption         =   "Items on &catalogue(s)"
         End
         Begin VB.Menu mnuWants 
            Caption         =   "&Wants"
         End
      End
      Begin VB.Menu mnuUtilities 
         Caption         =   "&Utilities"
         Begin VB.Menu mnuMerge 
            Caption         =   "&Merge two products"
         End
         Begin VB.Menu mnuMergeCust 
            Caption         =   "M&erge two customers/suppliers"
         End
         Begin VB.Menu mnuMergeCT 
            Caption         =   "Merge two& customer types"
         End
         Begin VB.Menu mnuMergePT 
            Caption         =   "Merge two product types"
         End
         Begin VB.Menu mnuMergeSEC 
            Caption         =   "Merge two sections"
         End
         Begin VB.Menu mnuMergeCurr 
            Caption         =   "Merge two currencies"
         End
      End
      Begin VB.Menu mnureloadconfiguration 
         Caption         =   "Reload configuration"
      End
      Begin VB.Menu mnuClearTemp 
         Caption         =   "Clear TEMP folder"
      End
      Begin VB.Menu mnuDiag 
         Caption         =   "&Diagnostics"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuControls 
      Caption         =   "&Reservations"
      Begin VB.Menu mnuReserve 
         Caption         =   "&On reserve list"
      End
   End
   Begin VB.Menu mnuMailing 
      Caption         =   "&CRM"
      Begin VB.Menu mnuCl 
         Caption         =   "&Customer Lists"
      End
      Begin VB.Menu mnuCustMail 
         Caption         =   "&Manage customer mailings"
      End
   End
   Begin VB.Menu mnuPOSale 
      Caption         =   "Point-of-Sale"
      Begin VB.Menu mnuCashup 
         Caption         =   "Cash-up"
      End
      Begin VB.Menu mnuCashupEx 
         Caption         =   "Cash-up by period"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuNonPOSSales 
         Caption         =   "Non POS sales (Invoices through Papyrus II)"
      End
      Begin VB.Menu mnuSPA 
         Caption         =   "&Sales person analysis"
      End
      Begin VB.Menu mnuOrderRequests 
         Caption         =   "&Order requests"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuFindForm 
      Caption         =   "FindForm"
      Visible         =   0   'False
      Begin VB.Menu mnuInventoryRecord 
         Caption         =   "Open inventory record"
      End
      Begin VB.Menu mnuPrintpickingSlip 
         Caption         =   "Print picking slip"
      End
      Begin VB.Menu mnuSOHAll 
         Caption         =   "Other branches' stock"
      End
      Begin VB.Menu mnuProductStatus 
         Caption         =   "Product status"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearTempList 
         Caption         =   "&Clear temporary list"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add to temporary list of products"
      End
      Begin VB.Menu mnuFFSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReorderFromBrowse 
         Caption         =   "Generate reorderslate"
      End
      Begin VB.Menu mnuCOPlace 
         Caption         =   "&Place on customer order"
      End
      Begin VB.Menu mnuQuickInvoice 
         Caption         =   "Create invoice"
      End
      Begin VB.Menu mnuPF 
         Caption         =   "Create pro-&forma"
      End
      Begin VB.Menu mnuLabels 
         Caption         =   "labels"
      End
      Begin VB.Menu mnuPlaceOnReserve 
         Caption         =   "Place on &reserve"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFFSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalesPatterns 
         Caption         =   "&SALES"
      End
      Begin VB.Menu mnuSpecialRequest 
         Caption         =   "Place on staff special request"
      End
      Begin VB.Menu mnuFFSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetPT 
         Caption         =   "Set product type"
      End
      Begin VB.Menu mnuSetSection 
         Caption         =   "Set category"
      End
      Begin VB.Menu mnuTouchRecord 
         Caption         =   "Send to P.O.S. computers"
      End
      Begin VB.Menu mnuMarkWeb 
         Caption         =   "Mark for Web export"
      End
   End
   Begin VB.Menu mnuCustomerBrowseContext 
      Caption         =   "CustomerBrowseContext"
      Visible         =   0   'False
      Begin VB.Menu mnuUpdatePOS_Cust 
         Caption         =   "Update P.O.S. computers"
      End
      Begin VB.Menu mnuAlertSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlert 
         Caption         =   "Alert to customer"
      End
      Begin VB.Menu mnuAlertHistory 
         Caption         =   "Show alert history"
      End
   End
   Begin VB.Menu mnuReorder 
      Caption         =   "REORDER"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveFromReorderList 
         Caption         =   "&Don't place supplier order"
      End
      Begin VB.Menu mnuSetDeal 
         Caption         =   "&Set supplier and deal"
      End
      Begin VB.Menu mnuReord_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalesPatterns2 
         Caption         =   "SALES"
      End
   End
   Begin VB.Menu mnuReserveList 
      Caption         =   "ReserveList"
      Visible         =   0   'False
      Begin VB.Menu mnuPutBack 
         Caption         =   "Put back into stock"
      End
      Begin VB.Menu mnuCustomerCollects 
         Caption         =   "Customer collects"
      End
   End
   Begin VB.Menu mnuReturnPopup 
      Caption         =   "ReturnPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRejected 
         Caption         =   "Alter rejected quantity"
      End
   End
   Begin VB.Menu mnuInvoicePreview 
      Caption         =   "InvoicePreview"
      Visible         =   0   'False
      Begin VB.Menu mnuSubstitute 
         Caption         =   "Insert substitute"
      End
      Begin VB.Menu mnuILine_COL 
         Caption         =   "View sales order line"
      End
   End
   Begin VB.Menu mnuShowOLHistGrp 
      Caption         =   "Previous versions"
      Visible         =   0   'False
      Begin VB.Menu mnuShowPOLHist 
         Caption         =   "Show previous versions"
      End
      Begin VB.Menu mnuPreDelAdv 
         Caption         =   "Pre-delivery advice message"
      End
      Begin VB.Menu mnuBrowseTActions 
         Caption         =   "Browse tracking actions"
      End
   End
   Begin VB.Menu mnPrevVerCO 
      Caption         =   "Previous versions"
      Visible         =   0   'False
      Begin VB.Menu mnuShowCOLHist 
         Caption         =   "Show previous versions"
      End
      Begin VB.Menu mnuBrowseTActionsCO 
         Caption         =   "Browse tracking actions"
      End
      Begin VB.Menu mnuDeliveryDoc 
         Caption         =   "Delivery document"
      End
   End
   Begin VB.Menu mnuBrowseInvoicesPopup 
      Caption         =   "Browse Invoices Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPrintSelected 
         Caption         =   "Print selected invoices"
      End
   End
   Begin VB.Menu mnuBrowseDeliveriesPopup 
      Caption         =   "Browse Deliveries Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPrintSelecteddeliveries 
         Caption         =   "Print selected deliveries"
      End
   End
   Begin VB.Menu mnuActionTransactionList 
      Caption         =   "ActionTransactionList"
      Visible         =   0   'False
      Begin VB.Menu mnuPrepareDetailList 
         Caption         =   "Prepare detail list for internet ordering"
      End
   End
   Begin VB.Menu mnuCustomerPreviewPopup 
      Caption         =   "CustomerPreviewPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuPayInvoice 
         Caption         =   "Pay this invoice"
      End
   End
   Begin VB.Menu mnuDelivery 
      Caption         =   "Delivery"
      Visible         =   0   'False
      Begin VB.Menu mnuDelClaim 
         Caption         =   "Claim"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuWeb 
         Caption         =   "Papyrus web site"
      End
      Begin VB.Menu mnuHelpdesk 
         Caption         =   "Helpdesk"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
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
    TOP             As Long
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

'Dim frmDebtors As frmDebtors
'Dim frmCreditors As frmCreditors
Dim frmBudgetPreview As frmBudgetPreview
Dim frmDashboard As frmDashboard
Dim frmConfiguration As frmConfiguration
Dim frmBrowseInvoices As frmBrowseInvoices
Dim frmBrowsePFInvoices As frmBrowseInvoices
Dim frmBrowseQuotations As frmBrowseQuotations
Dim frmBrowseCN As frmBrowseCN
'Dim frmBrowseJNL As frmBrowseDBJNLs
Dim frmBrowseCO As frmBrowseCOs
Dim frmBrowseCOsToInvoice As frmBrowseCOsToInvoice
Dim frmBrowseExchanges As frmBrowseExchanges
Dim frmBrowsePayment As frmBrowsePayments
Dim frmBrowsePO As frmBrowsePOs
Dim frmBrowseAPP As frmBrowseAPPs
Dim frmBrowseAPPR As frmBrowseAPPRs
Dim frmBrowseDEL As frmBrowseDels
Public frmBrowseSingles As frmBrowseSingles
Public frmBrowseProd As frmBrowseProducts
'Dim frmBrowseProdAQ As frmBrowseProductsAQ
Public frmREORDER_SAL As frmREORDER_CO
Public frmREORDER_CUST As frmREORDER_CO
Dim frmBrowseCustomers As frmBrowseCustomers
Public frmMainCustomerPreview As frmCustomerPreview
Public frmMainLoyaltyPreview As frmLoyaltyPreview
Public frmScratch As frmScratch
Public frmTRacking As frmTrackingActions
Public frmNewCust As frmNewCustomer
Public fTB As frmFindTextBite
Dim WithEvents oSQL As z_SQL
Attribute oSQL.VB_VarHelpID = -1

Private mlngPrevIndex As Long
Private Declare Function GetClassName Lib "user32" _
    Alias "GetClassNameA" (ByVal hWnd&, _
    ByVal lpClassName$, ByVal nMaxCount&) As Long

Dim bMouseDown As Boolean

Enum EnumMode
    eneditingrow = 0
    enAddingRow = 1
    enNotEditing = 3
End Enum
Dim bShiftDown As Boolean

Dim bForceClose As Boolean


Public Property Let ForceClose(bForce As Boolean)
    bForceClose = bForce
End Property


Private Sub MDIForm_Resize()
    'Budget
    If oPC.ShowBudget Then
        If frmBudgetPreview Is Nothing Then
            Set frmBudgetPreview = New frmBudgetPreview
            frmBudgetPreview.Hide
        End If
    
        frmBudgetPreview.TOP = Me.Height - Me.TOP - frmBudgetPreview.Height - 1500
        frmBudgetPreview.Left = 0
    End If
    
    'Dashboard
    If frmDashboard Is Nothing Then
        Set frmDashboard = New frmDashboard
    End If

    frmDashboard.TOP = Me.TOP + 10
    frmDashboard.Left = 0
End Sub

Private Sub mnuBrowsePF_Click()
    On Error GoTo errHandler
    BrowseInvoices True

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowsePF_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCreditors_Click()
'    If frmCreditors Is Nothing Then
'        Set frmCreditors = New frmCreditors
'    End If
'    frmCreditors.Show

End Sub

Private Sub mnuDebtors_Click()
'    If frmDebtors Is Nothing Then
'        Set frmDebtors = New frmDebtors
'    End If
'    frmDebtors.Show
End Sub




Private Sub mnuInventoryRecord_Click()
    frmBrowseProd.OpenInventoryRecord
End Sub

Private Sub mnuInvoicesToDeliver_Click()
Dim f As frmInvoicesToGDN
Dim oOF As z_OrderFulfilmentDocGen
Dim rs As ADODB.Recordset

    Set oOF = New z_OrderFulfilmentDocGen
    Set rs = oOF.CustomerDocsToDispatch
    If rs.eof Then
        MsgBox "There are no documents to dispatch.", vbInformation + vbOKOnly, "Status"
        rs.Close
        Set rs = Nothing
        Set oOF = Nothing
        Exit Sub
    End If
    Set f = New frmInvoicesToGDN
    f.component rs, True
    f.Show

End Sub



Private Sub mnuPreInv_Click()
    NewInvoice False, True

End Sub

Private Sub mnuPrintPickingSlip_Click()
    frmBrowseSingles.PrintPickingSlip
End Sub


Private Sub mnuAgedBalances_Click()
    On Error GoTo errHandler
Dim f As New frmAgedBalances
    f.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAgedBalances_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuAllSuppliers_Click()
    On Error GoTo errHandler
Dim frmR As New frmReturn1
Dim oSupp As a_Supplier


        
    frmR.component
    frmR.Show


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAllSuppliers_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub mnuBrowseBankPostings_Click()
'Dim f As New frmCashBook
'
'    f.Show
'
'End Sub

Private Sub mnuBrowseCategoryChecks_Click()
    On Error GoTo errHandler
Dim frm As New frmBrowseCategoryChecks

    frm.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseCategoryChecks_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuBrowseGoodsDelivery_Click()
Dim f As New frmBrowseGDNs

    f.Show
    
End Sub

Private Sub mnuBudget_Click()
Dim f As frmBudgetPreview

    If Not oPC.ShowBudget Then Exit Sub
    
    If frmBudgetPreview Is Nothing Then
        Set frmBudgetPreview = New frmBudgetPreview
    End If
        
    
    frmBudgetPreview.Visible = Not frmBudgetPreview.Visible
    
End Sub
Private Sub mnuDashboard_Click()

    
    If frmDashboard Is Nothing Then
        Set frmDashboard = New frmDashboard
    End If
        
    
    frmDashboard.ZOrder 0
    
End Sub


Private Sub mnuCOSR_Click()
    On Error GoTo errHandler
Dim frm As New frmBrowseCOSR

    frm.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCOSR_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuGDN_Click()
    NewGDN
End Sub

Private Sub mnuGDNsToInvoice_Click()
Dim f As frmGDNsToInvoice
Dim oOF As z_OrderFulfilmentDocGen
Dim rs As ADODB.Recordset

    Set oOF = New z_OrderFulfilmentDocGen
    Set rs = oOF.FetchOrderLinesToFulfil(False)
    If rs.eof Then
        MsgBox "There are no items to invoice.", vbInformation + vbOKOnly, "Status"
        rs.Close
        Set rs = Nothing
        Set oOF = Nothing
        Exit Sub
    End If
    Set f = New frmGDNsToInvoice
    f.component rs, True
    f.Show
    
End Sub
Private Sub mnuCompleteOrdersSnagged_Click()
Dim f As frmGDNsToInvoice
Dim oOF As z_OrderFulfilmentDocGen
Dim rs As ADODB.Recordset
    
    Set oOF = New z_OrderFulfilmentDocGen
    Set rs = oOF.FetchOrderLinesToFulfil(True)
    If rs.eof Then
        MsgBox "There are no items to list.", vbInformation + vbOKOnly, "Status"
        rs.Close
        Set rs = Nothing
        Set oOF = Nothing
        Exit Sub
    End If
    
    Set f = New frmGDNsToInvoice
    f.component rs, False
    f.Show

End Sub
Private Sub mnuImportBankfile_Click()
Dim f As New frmImportFromFile

    f.component enBankStatement
    f.Show

End Sub

Private Sub mnuImportForTransfer_Click()
    On Error GoTo errHandler
Dim frm As New frmImportTransferFile

    frm.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportForTransfer_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuInvoiceCompleteOrders_Click()
    If frmBrowseCOsToInvoice Is Nothing Then
       Set frmBrowseCOsToInvoice = New frmBrowseCOsToInvoice
    End If
    frmBrowseCOsToInvoice.ZOrder 0

End Sub

Private Sub mnuNewCategoryCheck_Click()
    On Error GoTo errHandler
Dim frm As frmCreateCategoryCheck

    Set frm = New frmCreateCategoryCheck
    frm.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewCategoryCheck_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDeliveryDoc_Click()
    On Error Resume Next
        Me.ActiveForm.mnuDeliveryDoc
End Sub

Private Sub mnuProductStatus_Click()
            On Error Resume Next
        
        Me.ActiveForm.mnuProductStatus

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuProductStatus_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuImportToClipboard_Click()
    On Error GoTo errHandler
Dim f As New frmImportFromFile
Dim rs As ADODB.Recordset
Dim rst As ADODB.Recordset
Dim oLine As a_POL
Dim fs As New FileSystemObject
Dim OpenResult As Integer

    f.component enClipboardImport
    f.Show vbModal
    Unload f
    oPC.PrepareLinesClipboard
    Set rs = oPC.LinesClipboard
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    rs.Open
    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM vLoadClipboard", oPC.COShort, adOpenDynamic
    Do While Not rst.eof
        rs.AddNew
        rs.Fields("GUID") = rst.Fields("GUID")
        rs.Fields("PID") = rst.Fields("PID")
        rs.Fields("Qty") = rst.Fields("QTY")
        rs.Fields("QtyFirm") = rst.Fields("QTYFIRM")
        rs.Fields("QtySS") = rst.Fields("QTYSS")
        rs.Fields("Price") = rst.Fields("PRICE")
        rs.Fields("DISCOUNTRATE") = rst.Fields("DiscountRate")
        rs.Fields("CODEF") = rst.Fields("CodeF")
        rs.Fields("EANF") = rst.Fields("EANF")
        rs.Fields("TITLE") = rst.Fields("Title")
        rs.Fields("VATRATE") = rst.Fields("VATRate")
        rs.Fields("REF") = rst.Fields("Ref")
        If Not IsNull(rst.Fields("ETA")) Then
            rs.Fields("ETA") = rst.Fields("ETA")
        End If
'        rs.Fields("EXTRACHARGEPID") = oLine.ExtraPID
'        rs.Fields("EXTRACHARGEVALUE") = oLine.ExtraCharge
        rst.MoveNext
        rs.Update
    Loop
    If Not fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        fs.CreateFolder (oPC.SharedFolderRoot & "\TEMP")
        If Err <> 0 Then
            MsgBox "Cannot create folder for Papyrus clipboard", vbInformation + vbOKOnly, "Can't do this"
        End If
    End If
    If fs.FileExists(oPC.SharedFolderRoot & "\TEMP\Clipboard.rs") Then
        fs.DeleteFile oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
    Else
        If fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
            rs.Save oPC.SharedFolderRoot & "\TEMP\Clipboard.rs"
        End If
    End If
    Set rs = Nothing
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuImportToClipboard_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportToClipboard_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPF_Click()
    On Error GoTo errHandler
    frmBrowseProd.PlacePF "PF"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPF_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuLabels_Click()
    On Error GoTo errHandler
    frmBrowseProd.PrintLabels
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuLabels_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuQuickInvoice_Click()
    On Error GoTo errHandler
    frmBrowseProd.PlacePF "INVOICE"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuQuickInvoice_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuHelpdesk_Click()
    On Error GoTo errHandler
Dim str As String
Dim str2 As String
    If oPC.InternetDialup = True Then Exit Sub
    Screen.MousePointer = vbHourglass
            str = "http://www.papyrussoftware.co.za/helpdesk"
            OpenBrowser str
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuHelpdesk_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub oSQL_ProgressB(lngPos As Long, lngMax As Long, pMsg As String)
    On Error GoTo errHandler
        Me.SB1.Panels(2).text = pMsg & CStr(lngPos) & " of " & CStr(lngMax)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.oSQL_ProgressB(lngPos,lngMax,pMsg)", Array(lngPos, lngMax, pMsg), EA_NORERAISE
    HandleError
End Sub

Private Sub NewServiceItem()
    On Error GoTo errHandler
Dim frm As frmServiceItem
Dim oProd As a_Product
    Set oProd = Constructor.CreateProduct(False)
    Set frm = New frmServiceItem
    frm.component oProd
    frm.Show
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewServiceItem"
End Sub
Private Sub ServiceItem()
    On Error GoTo errHandler
Dim frm As frmBrowseServiceItem
    Set frm = New frmBrowseServiceItem
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ServiceItem"
End Sub
Private Sub Catalogues()
    On Error GoTo errHandler
Dim frm As frmCatalogues
    Set frm = New frmCatalogues
    frm.Show
    Set frm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Catalogues"
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

Private Sub NewGenStock()
    On Error GoTo errHandler
Dim frmA As frmProductAQ
Dim frm As frmProductNB
Dim oProd As a_Product

    Set oProd = Constructor.CreateProduct(True)

    Set frm = New frmProductNB
    frm.component oProd
    frm.Show
    Set oProd = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewGenStock"
End Sub

Public Sub NewCustomer(pType As enumCustomerType)
    On Error GoTo errHandler
Dim frm As frmCustomer
Dim oCust As a_Customer
    Set frm = New frmCustomer
    Set oCust = New a_Customer
    oCust.BeginEdit
    oCust.InitializeNewCustomer pType
    frm.component oCust
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewCustomer(pType)", pType
End Sub
Public Sub NewLoyaltyCustomer()
    On Error GoTo errHandler
Dim frm As frmLoyalty
Dim oCust As a_Customer
    Set frm = New frmLoyalty
    Set oCust = New a_Customer
    oCust.BeginEdit
    oCust.InitializeNewCustomer enPrivate
    frm.component oCust
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewLoyaltyCustomer"
End Sub

Private Sub BrowseInvoices(bPF As Boolean)
    On Error GoTo errHandler
    If bPF Then
        If frmBrowsePFInvoices Is Nothing Then
           Set frmBrowsePFInvoices = New frmBrowseInvoices
        End If
        frmBrowsePFInvoices.component True
        frmBrowsePFInvoices.ZOrder 0
    Else
        If frmBrowseInvoices Is Nothing Then
           Set frmBrowseInvoices = New frmBrowseInvoices
        End If
        frmBrowseInvoices.component False
        frmBrowseInvoices.ZOrder 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseInvoices(bPF)", bPF
End Sub
Private Sub BrowseQuotes()
    On Error GoTo errHandler
    If frmBrowseQuotations Is Nothing Then
       Set frmBrowseQuotations = New frmBrowseQuotations
    End If
    frmBrowseQuotations.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseQuotes"
End Sub

Private Sub BrowseReturns()
    On Error GoTo errHandler
    If frmBrowseReturns Is Nothing Then
       Set frmBrowseReturns = New frmBrowseReturns
    End If
    frmBrowseReturns.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseReturns"
End Sub

Public Sub BrowseDELS()
    On Error GoTo errHandler
    If frmBrowseDEL Is Nothing Then
       Set frmBrowseDEL = New frmBrowseDels
    End If
    frmBrowseDEL.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseDELS"
End Sub
Private Sub BrowseTrans()
    On Error GoTo errHandler
    If frmBrowseTF Is Nothing Then
       Set frmBrowseTF = New frmBrowseTF
    End If
    frmBrowseTF.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseTrans"
End Sub
Public Sub BrowsePOs()
    On Error GoTo errHandler
    If frmBrowsePO Is Nothing Then
       Set frmBrowsePO = New frmBrowsePOs
    End If
    frmBrowsePO.ZOrder 0

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowsePOs"
End Sub
Public Sub BrowseOrders()
    On Error GoTo errHandler
    If frmBrowseCO Is Nothing Then
       Set frmBrowseCO = New frmBrowseCOs
    End If
    frmBrowseCO.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseOrders"
End Sub
Private Sub BrowseExchanges()
    On Error GoTo errHandler
    If frmBrowseExchanges Is Nothing Then
       Set frmBrowseExchanges = New frmBrowseExchanges
    End If
    frmBrowseExchanges.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseExchanges"
End Sub
Private Sub BrowsePayments()
    On Error GoTo errHandler
    If frmBrowsePayment Is Nothing Then
       Set frmBrowsePayment = New frmBrowsePayments
    End If
    frmBrowsePayment.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowsePayments"
End Sub

Private Sub BrowseCS()
    On Error GoTo errHandler
    If frmBrowseCS Is Nothing Then
       Set frmBrowseCS = New frmBrowseCS
    End If
    frmBrowseCS.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseCS"
End Sub
Private Sub BrowseCN()
    On Error GoTo errHandler
    If frmBrowseCN Is Nothing Then
       Set frmBrowseCN = New frmBrowseCN
    End If
    frmBrowseCN.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseCN"
End Sub
'Private Sub BrowseJNL()
'    On Error GoTo errHandler
'    If frmBrowseJNL Is Nothing Then
'       Set frmBrowseJNL = New frmBrowseDBJNLs
'    End If
'    frmBrowseJNL.ZOrder 0
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.BrowseJNL"
'End Sub
Private Sub BrowseApps()
    On Error GoTo errHandler
    If frmBrowseAPP Is Nothing Then
       Set frmBrowseAPP = New frmBrowseAPPs
    End If
    frmBrowseAPP.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseApps"
End Sub
Public Sub BrowseOrderRequests()
Dim frm As New frmOrderRequests
    frm.Show
End Sub
Private Sub BrowseAPPRs()
    On Error GoTo errHandler
    If frmBrowseAPPR Is Nothing Then
       Set frmBrowseAPPR = New frmBrowseAPPRs
    End If
    frmBrowseAPPR.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseAPPRs"
End Sub

Private Sub ShowTextBites()
    On Error GoTo errHandler
    If fTB Is Nothing Then
        Set fTB = New frmFindTextBite
    End If
    fTB.ZOrder 0
   fTB.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.ShowTextBites"
End Sub
Private Sub mnuAbout_Click()
    On Error GoTo errHandler
Dim frm As New frmAbout
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAbout_Click", , EA_NORERAISE
    HandleError
End Sub




'Private Sub mnuBrowseDBJNLS_Click()
'    On Error GoTo errHandler
'    BrowseJNL
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuBrowseDBJNLS_Click", , EA_NORERAISE
'    HandleError
'End Sub


Private Sub mnuBrowseTActionsCO_Click()
    On Error Resume Next
    Me.ActiveForm.BrowseTActions

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseTActionsCO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuBrowseTracking_Click()
    On Error GoTo errHandler
Dim oSQL As New z_SQL
'Dim rs As New ADODB.Recordset

'    Set rs = oSQL.GetTrackingActions("", "", 300)
    If frmTRacking Is Nothing Then
        Set frmTRacking = New frmTrackingActions
    End If
    frmTRacking.component "", ""  'rs
    frmTRacking.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseTracking_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuChangePeriods_Click()
    On Error GoTo errHandler
Dim frm As New frmPeriodSwitch
    frm.Show vbModal

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuChangePeriods_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCL_Click()
    On Error GoTo errHandler
Dim frm As frmBrowseCustomersEx
    Set frm = New frmBrowseCustomersEx
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCL_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuClearTemp_Click()
    On Error GoTo errHandler
Dim fs As New FileSystemObject
Dim fol As Object
Dim fils As Object
Dim f As File

    If MsgBox("You want to remove all files from the " & oPC.SharedFolderRoot & "\TEMP folder?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    If fs.FolderExists(oPC.SharedFolderRoot & "\TEMP") Then
        Set fils = fs.GetFolder(oPC.SharedFolderRoot & "\TEMP").Files
        For Each f In fils
            f.Delete
        Next
    End If
    If fs.FolderExists(oPC.SharedFolderRoot & "\PDF") Then
        Set fils = fs.GetFolder(oPC.SharedFolderRoot & "\PDF").Files
        For Each f In fils
            f.Delete
        Next
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuClearTemp_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub mnuDebtorJournal_Click()
    On Error GoTo errHandler
Dim frm1 As New frmCustJnl
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frm1 = New frmCustJnl
    frm1.component lngTPID, frm.CustomerName
    If Not bCancel Then
        frm1.Show vbModal
    Else
        Set frm1 = Nothing
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDebtorJournal_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuExchanges_Click()
    On Error GoTo errHandler
    BrowseExchanges
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExchanges_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub mnuDebtorsPayment_Click()
'    On Error GoTo errHandler
'Dim frm1 As frmCustPmt
'Dim lngTPID As Long
'Dim frm As frmBrowseCustomers2
'Dim bCancel As Boolean
'    Set frm = New frmBrowseCustomers2
'    frm.Show vbModal
'    lngTPID = frm.CustomerID
'    Unload frm
'    If lngTPID = 0 Then Exit Sub
'    Set frm1 = New frmCustPmt
'    frm1.component lngTPID, frm.CustomerName
'    If Not bCancel Then
'        frm1.Show vbModal
'    Else
'        Set frm1 = Nothing
'    End If
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuDebtorsPayment_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub mnuImportSales_Click()
    On Error GoTo errHandler
Dim frm As New frmSalesImport

    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuImportSales_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuMarkWeb_Click()
    On Error Resume Next
    Me.ActiveForm.SetForWebExport
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMarkWeb_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuMergeCurr_Click()
    On Error GoTo errHandler
Dim frm As New frmMergeCurrs

    frm.Show vbModal

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMergeCurr_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewAQBook_Click()
    On Error GoTo errHandler
Dim frmA As frmProductAQ
Dim oProd As a_Product
    
    Screen.MousePointer = vbHourglass
    Set oProd = Constructor.CreateProduct(True)
    Set frmA = New frmProductAQ
    frmA.component oProd
    frmA.Show
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewAQBook_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewBook_Click()
    On Error GoTo errHandler
Dim frm As frmProduct
Dim oProd As a_Product

'    Screen.MousePointer = vbHourglass
'    Set oProd = Constructor.CreateProduct(True)
'    Set frm = New frmProduct
'    frm.component oProd
'    frm.Show
'    Screen.MousePointer = vbDefault
    NewBook
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewBook_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub NewBook()
    On Error GoTo errHandler
Dim frmA As frmProductAQ
Dim frm As frmProduct
Dim oProd As a_Product
Dim frmST As frmSearchType
Dim strStyle As String
Dim frmS As frmProductSingles

    Set oProd = Constructor.CreateProduct(True)

        If CheckThisPoint(M_NEWPRODUCT) Then
            If SecurityControl(enSECURITY_CREATENEWSTOCKITEM, , "Creating new stock item", "You do not have permission to create new stock items (or your signature is invalid).") = False Then
                Exit Sub
            End If
        End If
    
    If oPC.AllowAntiquarionSearch = 1 Then
        Screen.MousePointer = vbHourglass
        If oPC.UniqueProducts Then
            Set frmS = New frmProductSingles
            frmS.component oProd
            frmS.Show
        Else
            Set frm = New frmProduct
            frm.component oProd
            frm.Show
        End If
        Screen.MousePointer = vbDefault
    ElseIf oPC.AllowAntiquarionSearch = 3 Then
        Screen.MousePointer = vbHourglass
        Set frmA = New frmProductAQ
        frmA.component oProd
        frmA.Show
        Screen.MousePointer = vbDefault
    ElseIf oPC.AllowAntiquarionSearch = 2 Then
        Set frmST = New frmSearchType
        frmST.Show vbModal
        strStyle = frmST.SearchType
        Unload frmST
        If strStyle = "N" Then
            Screen.MousePointer = vbHourglass
            Set frm = New frmProduct
            frm.component oProd
            frm.Show
            Screen.MousePointer = vbDefault
        Else
            Screen.MousePointer = vbHourglass
            Set frmA = New frmProductAQ
            frmA.component oProd
            frmA.Show
            Screen.MousePointer = vbDefault
        End If
    End If
    
     Set oProd = Nothing
   
    
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewBook"
End Sub
Private Sub BrowseBooks()

    On Error GoTo errHandler
Dim frm As frmSearchType
Dim strStyle As String

    If oPC.AllowAntiquarionSearch = 1 Then

        Screen.MousePointer = vbHourglass
            If oPC.UniqueProducts Then
                If frmBrowseSingles Is Nothing Then
                    Set frmBrowseSingles = New frmBrowseSingles
                    frmBrowseSingles.ZOrder 0
                End If
            Else
                If frmBrowseProd Is Nothing Then
                    Set frmBrowseProd = New frmBrowseProducts
                   ' MsgBox "Created"
                  '  MsgBox "frmbrowseProd is nothing" & (frmBrowseProd Is Nothing)
                End If
                frmBrowseProd.Show
                frmBrowseProd.ZOrder 0
           End If
        Screen.MousePointer = vbDefault
        
'170       ElseIf oPC.AllowAntiquarionSearch = 3 Then
'180           Screen.MousePointer = vbHourglass
'190           If frmBrowseProdAQ Is Nothing Then
'200               Set frmBrowseProdAQ = New frmBrowseProductsAQ
'210           End If
'220           frmBrowseProdAQ.ZOrder 0
'230           Screen.MousePointer = vbDefault
    ElseIf oPC.AllowAntiquarionSearch = 2 Then
            Screen.MousePointer = vbHourglass
            If frmBrowseProd Is Nothing Then
                Set frmBrowseProd = New frmBrowseProducts
            End If
            frmBrowseProd.Visible = True
            frmBrowseProd.ZOrder 0
            Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseBooks"

End Sub
'Private Sub BrowseAntBooks()
'    On Error GoTo errHandler
'    Screen.MousePointer = vbHourglass
'        If frmBrowseProdAQ Is Nothing Then
'            Set frmBrowseProdAQ = New frmBrowseProductsAQ
'        End If
'        frmBrowseProdAQ.ZOrder 0
'    Screen.MousePointer = vbDefault
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.BrowseAntBooks"
'End Sub
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
       Set frmBrowseCustomers = New frmBrowseCustomers
    End If
    frmBrowseCustomers.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseCustomers"
End Sub
Private Sub BrowseSupplier()
    On Error GoTo errHandler
    If frmBrowsesuppliers Is Nothing Then
       Set frmBrowsesuppliers = New frmBrowsesuppliers
    End If
    frmBrowsesuppliers.ZOrder 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.BrowseSupplier"
End Sub
Private Sub NewQuotation()
    On Error GoTo errHandler
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
Dim frmQ As frmQuotation
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    If frm.IsBlocked Then
        MsgBox "You cannot create a quotation for this customer as it is blocked", vbOKOnly, "Can't do this"
        Unload frm
        Exit Sub
    End If
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frmQ = New frmQuotation
    frmQ.component lngTPID
    If Not frmQ.Cancelled Then
        frmQ.Show
    Else
        Set frmQ = Nothing
    End If
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    Select Case Err
    Case Else
        MsgBox Error
        GoTo EXIT_Handler
        Resume
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewQuotation"
End Sub
Private Sub NewInvoice(Proforma As Boolean, PreDelivery As Boolean)
    On Error GoTo errHandler
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
Dim frmI As frmInvoice
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frmI = New frmInvoice
    frmI.component PreDelivery, lngTPID, , Proforma
    frmI.Show
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    Select Case Err
    Case Else
        MsgBox Error
        GoTo EXIT_Handler
        Resume
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewInvoice(Proforma)", Proforma
End Sub
Private Sub NewGDN()
    On Error GoTo errHandler
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
Dim frmG As frmGDN
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frmG = New frmGDN
    frmG.component lngTPID
    frmG.Show
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    Select Case Err
    Case Else
        MsgBox Error
        GoTo EXIT_Handler
        Resume
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewGDN"
End Sub

Private Sub NewCN()
    On Error GoTo errHandler
Dim frm1 As frmCN
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frm1 = New frmCN
    frm1.component lngTPID
    frm1.Show
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
ERR_Handler:
    Select Case Err
    Case Else
        MsgBox Error
        GoTo EXIT_Handler
        Resume
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewCN"
End Sub
Private Sub NewDEL()
    On Error GoTo errHandler
Dim frm2 As frmdelBB
Dim frmDelStyle2 As frmdel_Style2
Dim frmGRNWH As New frmGRNWH

Dim lngTPID As Long
Dim frm As frmBrowseSUppliers2
Dim bCancel As Boolean
    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    lngTPID = frm.SupplierID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    
    If oPC.UniqueProducts Then
        Set frmDelStyle2 = New frmdel_Style2
        frmDelStyle2.component bCancel, lngTPID
        If bCancel Then
            Unload frmDelStyle2
        Else
            frmDelStyle2.Show
        End If
    Else
        If oPC.DeliveryStyle = "WH" Then
            Set frmGRNWH = New frmGRNWH
          ' frmGRNWH.component bCancel, lngTPID
            If bCancel Then
                Unload frmGRNWH
            Else
                frmGRNWH.Show
            End If
        Else
            Set frm2 = New frmdelBB
            frm2.component bCancel, lngTPID
            If bCancel Then
                Unload frm2
            Else
                frm2.Show
            End If
        End If
    End If
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewDEL"
End Sub

Private Sub NewCO()
    On Error GoTo errHandler
Dim frm1 As frmCO
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    If lngTPID = 0 Then Exit Sub
    If frm.IsBlocked Then
        MsgBox "You cannot create a customer order for this customer as it is blocked", vbOKOnly, "Can't do this"
        Unload frm
        Exit Sub
    End If
    Unload frm
    Set frm1 = New frmCO
    frm1.component bCancel, , lngTPID
    If Not bCancel Then
        frm1.Show
    Else
        Set frm1 = Nothing
    End If
EXIT_Handler:
    Me.MousePointer = vbDefault
'ERR_Handler:
'    Select Case err
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
'        Resume
'    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewCO"
End Sub

Private Sub NewAPP()
    On Error GoTo errHandler
Dim frm1 As frmAPP
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    If frm.IsBlocked Then
        MsgBox "You cannot create an appro for this customer as it is blocked", vbOKOnly, "Can't do this"
        Unload frm
        Exit Sub
    End If
    
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frm1 = New frmAPP
    frm1.component , lngTPID
    frm1.Show
EXIT_Handler:
    Me.MousePointer = vbDefault
'ERR_Handler:
'    Select Case err
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
'        Resume
'    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewAPP"
End Sub
Private Sub NewAPPR()
    On Error GoTo errHandler
Dim frm1 As frmAPPR
Dim lngTPID As Long
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frm1 = New frmAPPR
    frm1.component , lngTPID
    frm1.Show
EXIT_Handler:
    Me.MousePointer = vbDefault
'ERR_Handler:
'    Select Case err
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
'        Resume
'    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewAPPR"
End Sub
Private Sub NewPO(pSubscriptionOrReplenishment As String)
    On Error GoTo errHandler
Dim frm1 As frmPO
Dim lngTPID As Long
Dim bCancel As Boolean
Dim frm As frmBrowseSUppliers2
    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    lngTPID = frm.SupplierID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frm1 = New frmPO
    frm1.component pSubscriptionOrReplenishment, bCancel, , lngTPID
    If bCancel Then
        Unload frm1
    Else
        frm1.Show
    End If
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewPO"
End Sub
Private Sub NewReturn()
    On Error GoTo errHandler
Dim frm1 As frmReturn
Dim lngTPID As Long
Dim bCancel As Boolean
Dim frm As frmBrowseSUppliers2
    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    lngTPID = frm.SupplierID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frm1 = New frmReturn
    frm1.component bCancel, , lngTPID
    If bCancel Then
        Unload frm1
    Else
        frm1.Show
    End If
EXIT_Handler:
    Me.MousePointer = vbDefault
'ERR_Handler:
'    Select Case err
'    Case Else
'        MsgBox Error
'        GoTo EXIT_Handler
      '  Resume
'    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewReturn"
End Sub


Private Sub MDIForm_Load()
    On Error GoTo errHandler
Dim strError As String
Dim cntTest As Long
Dim rs As ADODB.Recordset
Dim lkgPBKSDate As Date
Dim lkgMasterDate As Date
Dim s As String

    errSysHandlerSet

'MsgBox " D1"
    GetThunder
    PaintFirstScreen
  '  mnuStatements.Visible = oPC.RunsAccountsTF
'    If Not oPC.IsServerMachine Then
'        Me.mnuNewTestFromLive.Enabled = False
'        Me.mnuManDBCopies.Enabled = False
'
'    End If
'MsgBox " D2"
    mnuBrowseStock.Enabled = oPC.AllowGeneralStock
    mnuImportForTransfer.Enabled = (oPC.GetProperty("ServiceBroker_IBTs_ON") = "TRUE")
    mnuNBP.Enabled = oPC.AllowGeneralStock
    mnuCS.Visible = (oPC.GetProperty("ShowBrowseCashSalesMenu") = "TRUE")
    TBHEAD.Buttons("bcs").Visible = (oPC.GetProperty("ShowBrowseCashSalesMenu") = "TRUE")
    mnuTFRIn.Visible = (oPC.GetProperty("ShowTransferInMenu") = "TRUE")
    TBHEAD.Buttons("bGen").Visible = oPC.AllowGeneralStock
    TBHEAD.Buttons("ngs").Visible = oPC.AllowGeneralStock
    mnuCOSR.Visible = oPC.SupportsUNISA
'MsgBox " D3"
    
    mnuPrintpickingSlip.Visible = oPC.UniqueProducts
    mnuReorderFromBrowse.Visible = Not oPC.UniqueProducts
    mnuOrderFulfilmentAppros.Visible = oPC.CanGenerateApprosFromOrderFulfilment
    mnuNewInvoice.Caption = IIf(oPC.GetProperty("AllowsGDNs") = "TRUE", "Counter invoice", "Invoice")
    If oPC.GetProperty("AllowsGDNs") = "TRUE" Then
        mnuPreInv.Visible = True
        Me.TBHEAD.Buttons(13).ToolTipText = "New counter invoice"
        mnuNewInvoice.Caption = "Counter invoice"
        mnuPastelinestoNewInvoice.Caption = "counter invoice"
    Else
        mnuPreInv.Visible = False
        Me.TBHEAD.Buttons(13).ToolTipText = "New invoice"
        mnuNewInvoice.Caption = "Invoice"
        mnuPastelinestoNewInvoice.Caption = "Invoice"
        mnuPastelinestoNewCounterInvoice.Caption = "Invoice"
        mnuGDN.Visible = False
    End If
    If Not fRunningInIde Then
        subclassMDIClientArea Me
        DrawLogo GetProp(Me.hWnd, "MAINhMDIClient")
    End If
'MsgBox " D4"

    cntTest = 0
    If UCase(oPC.GetProperty("TestDatabase")) = "TRUE" Then
        oPC.OpenDBSHort
restart:
        Set rs = New ADODB.Recordset
        rs.Open "SELECT LNG_LastKnownGood FROM tLastKnownGood WHERE LNG_Databasename = '" & oPC.DatabaseName & "'", oPC.COShort, adOpenStatic
        If Not (rs.eof And rs.BOF) Then
            lkgPBKSDate = rs.Fields(0)
            rs.Close
            Set rs = Nothing
            Set rs = New ADODB.Recordset
            rs.Open "SELECT LNG_LastKnownGood FROM tLastKnownGood WHERE LNG_Databasename = 'Master'", oPC.COShort, adOpenStatic
            lkgMasterDate = rs.Fields(0)
            rs.Close
            Set rs = Nothing
        End If
        If (DateDiff("d", lkgMasterDate, Date) > 1 Or DateDiff("d", lkgPBKSDate, Date) > 1) And (oPC.DatabaseName = "PBKS") Then
            If cntTest = 0 Then
                cntTest = cntTest + 1
                MsgBox "Papyrus needs to run a database check and this may take up to 5 minutes." & vbCrLf & "Inform other users to wait until this completes before starting Papyrus applications." & vbCrLf & "Click OK to begin checking.", vbOKOnly, "Database needs checking."
                WaitMsg "Checking database, please wait", True
                Dim tmp As String
                tmp = SB1.Panels(2).text
                SB1.Panels(2).text = "Running check on database. Wait . . ."
                oPC.COShort.CommandTimeout = 0
                oPC.COShort.execute " DBCC CHECKDB ('" & oPC.DatabaseName & "') WITH NO_INFOMSGS;"
                oPC.COShort.execute " DBCC CHECKDB ('Master') WITH NO_INFOMSGS;"
                oPC.COShort.execute "dbo.UpdateLNG ('" & oPC.DatabaseName & "')"
                SB1.Panels(2).text = tmp
                WaitMsg "Checking database, please wait", False

                GoTo restart
            End If
            If DateDiff("d", lkgMasterDate, Date) > 1 Then
                s = "It looks like the Master database may be damaged."
            End If
            If DateDiff("d", lkgPBKSDate, Date) > 1 Then
                s = s & "It looks like the Papyrus database may be damaged."
            End If
        Else
            If cntTest > 0 Then
                s = "It looks like the dayend process did not run last night."
            End If
        End If

        
        If s > "" Then
            MsgBox s & vbCrLf & "If the server was off last night this is expected, otherwise please contact your support person.", vbInformation + vbOKOnly, "Important warning"
        End If
        
        oPC.DisconnectDBShort
    End If

'MsgBox " D10"





    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub PaintFirstScreen()
    On Error GoTo errHandler
    Me.BackColor = RGB(36, 60, 140)
      Dim fs As New FileSystemObject
    If UBound(arCommandLine()) > 0 Then
        If arCommandLine(1) <> "N" Then
            BackColor = vbRed
        End If
    End If
    If oPC.DatabaseName = "PBKS_TEST" Then
        BackColor = vbRed
    End If
    If oPC.BFLoaded Then
        Me.Caption = "Papyrus II Manager v1.3  -   Connected to Bookfind"
    Else
        Me.Caption = "Papyrus II Manager v1.3"
    End If
    If oPC.POSActive Then
        Me.Caption = Me.Caption & "    Point-of-sale active"
    End If
    Set fs = CreateObject("Scripting.FileSystemObject")
  If oPC.IsServerMachine Then
        On Error Resume Next
        Dim s As String
        s = ""
        If fs.FileExists("c:\PBKS\BU\PBKS.BAK") Then
            Dim x As File
            Set x = fs.GetFile("c:\PBKS\BU\PBKS.BAK")
            s = CStr(x.DateLastModified)
    
            If DateDiff("d", x.DateLastModified, Date) > 2 Then
                Me.SB1.Panels(1) = "Last day-end: " & oPC.Configuration.LastUpdateDateF & " ******* Date of backup:" & s & " ********"
            Else
                Me.SB1.Panels(1) = "Last day-end: " & oPC.Configuration.LastUpdateDateF & "  Date of backup:" & s
            End If
            On Error GoTo errHandler
        Else
            Me.SB1.Panels(1) = "Last day-end: " & oPC.Configuration.LastUpdateDateF & " ******* Date of backup:" & s & " ********"

        End If
    Else
        Me.SB1.Panels(1) = "Last day-end: " & oPC.Configuration.LastUpdateDateF
    End If
  
    Me.SB1.Panels(2) = " " & oPC.NewQuotation
    Me.SB1.Panels(3) = " " & IIf(oPC.DatabaseName <> "PBKS", "Server:" & oPC.servername & ", Database:" & oPC.DatabaseName, "Server:" & oPC.servername)
    SB1.Panels(2).ToolTipText = SB1.Panels(2).text
    mnuPOSale.Visible = oPC.POSActive
    
    Select Case oPC.AllowAntiquarionSearch
    Case 1   'New books only
        mnuNewBook.Visible = True
        mnuBrowseBooks.Visible = True
      '  mnuNewAQBook.Visible = False
      '  mnuBrowseAQBooks.Visible = False
    Case 2  'Both
        mnuNewBook.Visible = True
        mnuBrowseBooks.Visible = True
      '  mnuNewAQBook.Visible = True
      '  mnuBrowseAQBooks.Visible = True
    Case 3    'Antiquarian only
        mnuNewBook.Visible = False
        mnuBrowseBooks.Visible = False
     '   mnuNewAQBook.Visible = True
     '   mnuBrowseAQBooks.Visible = True
    End Select
    mnuFD.Visible = (oPC.POSActive = True)
'''''''''''''''  '  mnuBrowseDBJNLS.Visible = oPC.RunsAccountsTF
'''''''''''''''   ' mnubrowsePayments.Visible = oPC.RunsAccountsTF
'''''''''''''''  '  mnuDebtorJournal.Visible = oPC.RunsAccountsTF
'''''''''''''''  '  mnuDebtorsPayment.Visible = oPC.RunsAccountsTF
'    mnuNewCustomer_BC.Visible = oPC.SupportsBookClubsTF
'    mnuNewLoyalty.Visible = oPC.SupportsLoyaltyCustomersTF
    
    mnuBrowseGoodsDelivery.Visible = oPC.IncludeSupplierFeatures
    mnuGDN.Visible = oPC.IncludeSupplierFeatures
    
    mnuOF.Visible = oPC.IncludeSupplierFeatures
    
    hx.Visible = oPC.RunsAccountsTF
    h4b.Visible = oPC.RunsAccountsTF
    
    mnuPOSale.Visible = oPC.POSActive
    
    mnuFD.Visible = oPC.POSActive
    mnuControls.Visible = False 'oPC.POSActive
'    If Not fRunningInIde And oPC.DatabaseName <> "PBKS_TEST" Then
'        subclassMDIClientArea Me
'        DrawLogo GetProp(Me.hwnd, "MAINhMDIClient")
'    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PaintFirstScreen"
End Sub
Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    If Not bForceClose Then
        If UnloadMode = 0 Or UnloadMode = 1 Then
            If MsgBox("You want to close Papyrus II manager?", vbQuestion + vbYesNo, "Application closing") = vbNo Then
                Cancel = True
            End If
    End If
'        LogSaveToFile "Form count = " & Forms.Count
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode), EA_NORERAISE
    HandleError
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error GoTo errHandler
Dim i As Integer

    If oPC.ShowBudget Then Unload frmBudgetPreview
    Unload frmDashboard
    Me.SB1.Panels(2).text = "Cleaning up temporary files . . ."
    unsubclassMDIClientArea Me
   ' CleanupOldFilesInTempFolders
    Set frmMain = Nothing
    Set oPC = Nothing
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.MDIForm_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAdjust_Click()
    On Error Resume Next
    Me.ActiveForm.mnuAdjust
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAdjust_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuAPP_Click()
    On Error GoTo errHandler
    NewAPP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAPP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAPPR_Click()
    On Error GoTo errHandler
    NewAPPR
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAPPR_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAPPRs_Click()
    On Error GoTo errHandler
    BrowseAPPRs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAPPRs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuApps_Click()
    On Error GoTo errHandler
    BrowseApps
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuApps_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuBIC_Click()
    On Error GoTo errHandler
'Dim frm As New frmBICImport
'    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBIC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuBrowseBooks_Click()
    On Error GoTo errHandler
    BrowseBooks
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseBooks_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub mnuBrowseAQBooks_Click()
'    On Error GoTo errHandler
'    BrowseAntBooks
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuBrowseAQBooks_Click", , EA_NORERAISE
'    HandleError
'End Sub
Private Sub mnuBrowseCustomers_Click()
    On Error GoTo errHandler
    BrowseCustomers
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseCustomers_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuBrowseInvoices_Click()
    On Error GoTo errHandler
    BrowseInvoices False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseInvoices_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuBrowseQuotes_Click()
    On Error GoTo errHandler
    BrowseQuotes
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseQuotes_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuBrowseStock_Click()
    On Error GoTo errHandler
    BrowseGenStock
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseStock_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuCashbookDebtorSelect_Click()
    On Error Resume Next
    Me.ActiveForm.mnuSelectDebtor
End Sub
Private Sub mnuPaymentMatch_Click()
    On Error Resume Next
    Me.ActiveForm.mnuPaymentMatch
End Sub
Private Sub mnuBrowseSUpp_Click()
    On Error GoTo errHandler
    BrowseSupplier
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseSUpp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCancel_Click()
    On Error Resume Next
    Me.ActiveForm.mnuCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCancelINactive_Click()
    On Error Resume Next
    Me.ActiveForm.mnuCancelInactiveLines
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCancelINactive_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCancelLine_Click()
    On Error Resume Next
    Me.ActiveForm.mnuCancelLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCancelLine_Click", , EA_NORERAISE, , "FormName", Array(Me.ActiveForm.Name)
    HandleError
End Sub

Private Sub mnuCashup_Click()
    On Error GoTo errHandler
Dim frm As frmCashUP
Dim f As frmCashUPForBlind
Dim errRepeat As Integer

    errSysHandlerSet
    
    errRepeat = 0
   
    If SecurityControl(enSECURITY_CASHUP_SIGN, , "Opening cash-up forms", "You do not have permission to open the cash-up forms.") = False Then Exit Sub
    If oPC.BlindCashup Then
        Set f = New frmCashUPForBlind
        f.Show
    Else
        Set frm = New frmCashUP
        frm.Show
    End If
    Exit Sub
errHandler:

'new 23/1/2012
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmMain: mnuCashup_Click, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmMain: mnuCashup_Click after 5 re-attempts"
            MsgBox "Memory error in mnuCashup_Click. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
'''''''

    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCashup_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuChangePassword_Click()
    On Error GoTo errHandler
Dim frm As frmPWDChange

    Set frm = New frmPWDChange
    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuChangePassword_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCNote_Click()
    On Error GoTo errHandler
    NewCN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCNote_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuFD_Click()
    On Error GoTo errHandler
Dim frm As frmExchanges1
    Set frm = New frmExchanges1
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuFD_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuFulfil_Click()
    On Error Resume Next
    Me.ActiveForm.mnuFulfilLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuFulfil_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCatalogues_Click()
    On Error GoTo errHandler
    Catalogues
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCatalogues_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCathead_Click()
    On Error GoTo errHandler
Dim frm As frmCathead
    Set frm = New frmCathead
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCathead_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuCNotes_Click()
    On Error GoTo errHandler
    BrowseCN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCNotes_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCO_Click()
    On Error GoTo errHandler
    NewCO
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuConfig_Click()
    On Error GoTo errHandler
    
    If frmConfiguration Is Nothing Then
       Set frmConfiguration = New frmConfiguration
    End If
    
    If SecurityControl(enSECURITY_CONFIG_SIGN, , "Entering configuration", "You do not have permission to open the configuration form.") = False Then Exit Sub
    Set frmConfiguration = New frmConfiguration
    frmConfiguration.component oPC.Configuration
    frmConfiguration.ZOrder 0
    
    frmConfiguration.Show 'vbModal
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuConfig_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuConfig_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCOOD_Click()
    On Error GoTo errHandler
Dim frm As New frmLoadODCO
    frm.component "CO"
    frm.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCOOD_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCOs_Click()
    On Error GoTo errHandler
    BrowseOrders
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCOs_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnubrowsePayments_Click()
    On Error GoTo errHandler
    BrowsePayments
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnubrowsePayments_Click", , EA_NORERAISE
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
    ErrorIn "frmMain.mnuCountries_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCS_Click()
    On Error GoTo errHandler
    BrowseCS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCS_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCustMail_Click()
    On Error GoTo errHandler
Dim frmMail As New frmMailing
    frmMail.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCustMail_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDel_Click()
    On Error GoTo errHandler
    NewDEL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDelLine_Click()
    On Error Resume Next
    Me.ActiveForm.mnuDelLine
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDelLine_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuCreateCreditNote_Click()
    On Error Resume Next
    Me.ActiveForm.CreateCreditNote
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCreateCreditNote_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuDels_Click()
    On Error GoTo errHandler
    BrowseDELS
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDels_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDiag_Click()
    On Error GoTo errHandler
Dim frm As New frmDiagnostics
    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDiag_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDictionary_Click()
    On Error GoTo errHandler
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
'        MsgBox "You do not have security to edit dictionary.", vbExclamation, "Denied"
'        Exit Sub
'    End If
    If SecurityControl(enSECURITY_DICT_SIGN, , "Enter your security code.", "You do not have permission to open the dictionary.") = False Then Exit Sub

    Dictionary
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDictionary_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuexCapt_Click()
    On Error GoTo errHandler
Dim frm As frmCapturedSince
    Set frm = New frmCapturedSince
    frm.Caption = "Create list of all books captured since"
    Set frm = New frmCapturedSince
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuexCapt_Click", , EA_NORERAISE
    HandleError
End Sub


'Private Sub mnuDmpCUST_Click()
'Dim frm As New frmExportCUST
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
'        MsgBox "You do not have security to export records.", vbExclamation, "Denied"
'        Exit Sub
'    End If
'    frm.Show vbModal
'End Sub
'
'Private Sub mnuDmpPROD_Click()
'Dim frm As New frmExportPROD
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
'        MsgBox "You do not have security to export records.", vbExclamation, "Denied"
'        Exit Sub
'    End If
'    frm.Show vbModal
'End Sub

'Private Sub mnuDmpSUPP_Click()
'Dim frm As New frmExportSUPP
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
'        MsgBox "You do not have security to export records.", vbExclamation, "Denied"
'        Exit Sub
'    End If
'    frm.Show vbModal
'End Sub
'Private Sub mnuDmpDocs_Click()
'Dim frm As New frmExportTR
'Dim frmS As New frmSecurity
'    frmS.Show vbModal
'    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
'        MsgBox "You do not have security to export records.", vbExclamation, "Denied"
'        Exit Sub
'    End If
'    frm.Show vbModal
'
'End Sub

Private Sub mnuEXCat_Click()
    On Error GoTo errHandler
Dim frm As frmPSDetails
    Set frm = New frmPSDetails
    frm.Caption = "Create list of all books on one or many catalogue(s)"
    frm.Show vbModal
    Unload frm
    Set frm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuEXCat_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuExit_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExit_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuNewCustomer_BC_Click()
    On Error GoTo errHandler
    NewCustomer enBookclub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewCustomer_BC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewCustomer_Bus_Click()
    On Error GoTo errHandler
    NewCustomer enBusiness
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewCustomer_Bus_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewTestFromLive_Click()
    On Error GoTo errHandler
Dim oBU As z_PBKSBackup
Dim fs As New FileSystemObject
Dim strFilefolder As String
Dim strFilename As String
Dim tmp As String

    tmp = Me.SB1.Panels(2).text
    Me.SB1.Panels(2).text = "copying database . . . "
    strFilename = oPC.SharedFolderRoot & "\BU\PBKS_TEST.BAK"
    
    Set oBU = New z_PBKSBackup
    Screen.MousePointer = vbHourglass
    
    oBU.BackupToBriefcase strFilename, True, True
            DoEvents
    Me.SB1.Panels(2).text = "Attaching new test database . . . "
    
    Screen.MousePointer = vbDefault
    MsgBox "New test database has been created. You are still connected to the " & IIf(oPC.DatabaseName = "PBKS_TEST", "TEST", "LIVE") & " database", vbOKOnly, "Status"
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewTestFromLive_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuBackupCur_Click()
    On Error GoTo errHandler
Dim oBU As z_PBKSBackup
Dim fs As New FileSystemObject
Dim strFilefolder As String
Dim strFilename As String
Dim tmp As String


    CD1.DialogTitle = "Save to"
    CD1.DefaultExt = "BAK"
    CD1.InitDir = oPC.BackupFolder
    CD1.FLAGS = cdlOFNOverwritePrompt
    CD1.ShowSave
    strFilename = CD1.FileName
    
    tmp = Me.SB1.Panels(2).text
    Me.SB1.Panels(2).text = "copying database . . . "
    
    Set oBU = New z_PBKSBackup
    Screen.MousePointer = vbHourglass
    
    oBU.BackupToBriefcase strFilename, False, True
            DoEvents
    SB1.Panels(2).text = oPC.NewQuotation
    
    Screen.MousePointer = vbDefault
    MsgBox "New test database has been created. You are still connected to the " & IIf(oPC.DatabaseName = "PBKS_TEST", "TEST", "LIVE") & " database", vbOKOnly, "Status"
    SB1.Panels(2).text = oPC.NewQuotation


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBackupCur_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewTransfer_Click()
    On Error GoTo errHandler
Dim frm1 As frmTFPreview
Dim lngTPID As Long
Dim frm As frmBrowseStores
Dim bCancel As Boolean
Dim strStorename As String
Dim strFilename As String
Dim lngTRID As Long
    Set frm = New frmBrowseStores
    frm.Show vbModal
    lngTPID = frm.StoreID
    strStorename = frm.StoreName
    Unload frm
    If lngTPID = 0 Then Exit Sub
    
    CD1.DialogTitle = "Find XML file containing transfer details"
    CD1.DefaultExt = ".XML"
    CD1.InitDir = "C:\PBKS"
    CD1.ShowOpen
    strFilename = CD1.FileName
    CreateTransfer_IN strFilename, lngTPID, lngTRID
    
    Set frm1 = New frmTFPreview
    frm1.component lngTRID
    frm1.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewTransfer_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub CreateTransfer_IN(pFilename As String, pStoreID As Long, pNewID As Long)
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.CreateTransferFromXML pStoreID, pFilename, pNewID
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.CreateTransfer_IN(pFilename,pStoreID,pNewID)", Array(pFilename, pStoreID, _
         pNewID), EA_NORERAISE
    HandleError
End Sub
Private Sub mnuNonPOSSales_Click()
    On Error GoTo errHandler
Dim frmDlg As New frmNonPOSSales

    frmDlg.Show 'vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNonPOSSales_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub mnuODCOStatus_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
    oSM.PrintCustomerOrderStatusReport oPC.WorkstationName

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuODCOStatus_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub mnuPrepareStatements_Click()
    On Error GoTo errHandler
Dim oStatements As New a_Statements

    If MsgBox("Confirm you want to prepare the statements." & vbCrLf & "This could take a few minutes.", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Set oStatements = New a_Statements
    
    
    oStatements.PrepareStatements oPC.GetProperty("OnlyActiveAccounts") = "TRUE"
    Set oSQL = Nothing
    
    Screen.MousePointer = vbDefault
    MsgBox "Statement files have been created. They are ready to print", vbInformation, "Status"
    SB1.Panels(2).text = oPC.NewQuotation

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPrepareStatements_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPrintStatements_Click()
    On Error GoTo errHandler
Dim oFS As New FileSystemObject
Dim oXML As zXML
Dim fol, fil, f
    If MsgBox("Confirm you want to print the statements.", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Me.SB1.Panels(2) = "Statements are being printed . . . "
    DoEvents
    Set fol = oFS.GetFolder(oPC.SharedFolderRoot & "\Statements\")
    Set fil = fol.Files
    For Each f In fil
        Set oXML = New zXML
        oXML.PrintXML oPC.SharedFolderRoot & "\Statements\" & f.Name, oPC.SharedFolderRoot & "\TEMP", oPC.SharedFolderRoot & "\Templates\", oPC.LocalFolder & "\Executables", True
        Set oXML = Nothing
    Next
    Screen.MousePointer = vbDefault
    SB1.Panels(2).text = oPC.NewQuotation
    MsgBox "The statements have finished printing.", vbInformation, "Status"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPrintStatements_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuRemittance_Click()
    On Error Resume Next
    Me.ActiveForm.mnuLoadRemittance
End Sub

Private Sub mnuRestoreTest_Click()
    On Error GoTo errHandler
Dim oDMO As New z_SQLDMO
Dim strFilename As String
    CD1.DialogTitle = "Find .BAK file to restore"
    CD1.DefaultExt = "BAK"
    CD1.InitDir = oPC.BackupFolder
    CD1.CancelError = True
    CD1.ShowOpen
    
    Screen.MousePointer = vbHourglass
    strFilename = CD1.FileName
    oDMO.RestoreDatabase strFilename
    
    Screen.MousePointer = vbDefault
    MsgBox "The TEST database has been created from the backup file: " & strFilename & ". However you are still working on the " & IIf(oPC.DatabaseName = "PBKS_TEST", "TEST", "LIVE") & " database.", vbOKOnly, "Status"
    
    
CANCELERROR_ROUTINE:
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRestoreTest_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuRevPapClip_Click()
    On Error GoTo errHandler

    If oPC.QtyLinesinClipboard = 0 Then
        MsgBox "The Papyrus clipboard is empty.", vbInformation, "Can't do this"
        Exit Sub
    End If
    If frmScratch Is Nothing Then
       Set frmScratch = New frmScratch
    End If
    frmScratch.ZOrder 0

    frmScratch.component oPC.LinesClipboard
    frmScratch.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRevPapClip_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSalesComm_Click()
    On Error Resume Next
    Me.ActiveForm.mnuSalesComm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSalesComm_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSelSupplier_Click()
    On Error GoTo errHandler
Dim frm As New frmBrowseSUppliers2
Dim frmR As New frmReturn1
Dim oSupp As a_Supplier


    frm.Caption = "Select supplier for return"
    frm.Show vbModal
    If frm.SupplierID > 0 Then
        Set oSupp = New a_Supplier
        oSupp.Load frm.SupplierID
    End If
    Unload frm
    If oSupp Is Nothing Then Exit Sub
    frmR.component oSupp
    frmR.Show


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSelSupplier_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuMemo_Click()
    On Error Resume Next
    Me.ActiveForm.mnuMemo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMemo_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuMerge_Click()
    On Error GoTo errHandler
Dim frm As frmMergeProducts
Dim frmS As New frmSecurity

    If SecurityControl(enSECURITY_ISSUPERVISOR, , "Merge products", "You do not have authority to merge products.") = False Then
           Exit Sub
    End If

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
    ErrorIn "frmMain.mnuMerge_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuSOHAll_Click()
            On Error Resume Next
    Me.ActiveForm.mnuFindAllSOH

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSOHAll_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuDelClaim_Click()
    
    On Error Resume Next
    Me.ActiveForm.mnuOpenClaim
    
End Sub
Private Sub mnuMergeCust_Click()
    On Error GoTo errHandler
Dim frm As frmMergeTPs
Dim frmS As New frmSecurity
    frmS.Show vbModal
    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
        MsgBox "You do not have security to merge customers or suppliers.", vbExclamation, "Denied"
        Exit Sub
    End If
    Set frm = New frmMergeTPs
    frm.Show vbModal
    Unload frm
    Set frm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMergeCust_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuDiscr_Click()
    On Error Resume Next
    Me.ActiveForm.PrintSupplierClaim
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuDiscr_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuBarcodes_Click()
    On Error Resume Next
    Me.ActiveForm.PrintLabels
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBarcodes_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuCustomerAlloc_Click()
    On Error Resume Next
    Me.ActiveForm.CustomerAllocations
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCustomerAlloc_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuMergePT_Click()
    On Error GoTo errHandler
Dim frm As New frmMergePTs
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMergePT_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuMergeSEC_Click()
    On Error GoTo errHandler
Dim frm As New frmMergeSECs
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMergeSEC_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuMergeCT_Click()
    On Error GoTo errHandler
Dim frm As New frmMergeCustomerTypes
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuMergeCT_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuNBP_Click()
    On Error GoTo errHandler
    NewGenStock
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNBP_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub mnuNewCustomer_Click()
    On Error GoTo errHandler
    Set frmNewCust = New frmNewCustomer
    frmNewCust.Show
 '   NewCustomer enPrivate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewCustomer_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewInvoice_Click()
    On Error GoTo errHandler
    NewInvoice False, False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewInvoice_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuQuote_Click()
    On Error GoTo errHandler
    NewQuotation
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuQuote_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuNewLoyalty_Click()
    On Error GoTo errHandler
    NewLoyaltyCustomer
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewLoyalty_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewReturn_Click()
    On Error GoTo errHandler
    NewReturn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewReturn_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub NewSupplier()
    On Error GoTo errHandler
Dim frm As frmSupplier
Dim oSupp As a_Supplier
    Set frm = New frmSupplier
    Set oSupp = New a_Supplier
    frm.component oSupp
    frm.Show
    Exit Sub
errHandler:
    ErrorIn "frmMain.NewSupplier"
End Sub





Private Sub mnuNewStock_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewStock_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewSupp_Click()
    On Error GoTo errHandler
    NewSupplier
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewSupp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuNewServiceItem_Click()
    On Error GoTo errHandler
    NewServiceItem
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuNewServiceItem_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuServiceItem_Click()
    On Error GoTo errHandler
    ServiceItem
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuServiceItem_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuODPO_Click()
    On Error GoTo errHandler
Dim frm As New frmLoadODPO
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuODPO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuOrderRequests_Click()
    On Error GoTo errHandler
    BrowseOrderRequests
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuOrderRequests_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPlaceOnReserve_Click()
    On Error Resume Next
    Me.ActiveForm.PlaceOnReserve
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPlaceOnReserve_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPO_Click()
    On Error GoTo errHandler
    NewPO "R"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPOs_Click()
    On Error GoTo errHandler
    BrowsePOs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPOs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuProforma_Click()
    On Error GoTo errHandler
    NewInvoice True, False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuProforma_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPT_Click()
    On Error GoTo errHandler
Dim frm As frmPTs
Dim frmS As New frmSecurity
    frmS.Show vbModal
    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
        MsgBox "You do not have security to edit product types.", vbExclamation, "Denied"
        Exit Sub
    End If

    Set frm = New frmPTs
    frm.Show 'vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPT_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPurch_Click()
    On Error GoTo errHandler

    Set frmREORDER_CUST = New frmREORDER_CO
    frmREORDER_CUST.component "CUST"
    frmREORDER_CUST.Show
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPurch_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnureloadconfiguration_Click()
    On Error GoTo errHandler
    oPC.ReloadConfiguration
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnureloadconfiguration_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuReminders_Click()
    On Error GoTo errHandler
Dim oSM As New z_StockManager
Dim frm As frmPrintRemindersheet

        Screen.MousePointer = vbDefault
        Set frm = New frmPrintRemindersheet
        frm.Show vbModal
        oSM.PrintPurchaseOrderReminderReport oPC.WorkstationName, frm.chkPagePerSupplier = 1
        Unload frm
        Screen.MousePointer = vbDefault
   
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReminders_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuREORDSales_Click()
    On Error GoTo errHandler
    Set frmREORDER_SAL = New frmREORDER_CO
    frmREORDER_SAL.component "SALES"   ', lngTPID
    frmREORDER_SAL.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuREORDSales_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuReserve_Click()
    On Error GoTo errHandler
Dim frm As frmReserveList
    Set frm = New frmReserveList
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReserve_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuReturns_Click()
    On Error GoTo errHandler
    BrowseReturns
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReturns_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuRR_Click()
    On Error GoTo errHandler
Dim frm As frmRR
Dim frmS As New frmSecurity
    frmS.Show vbModal
    If oPC.Configuration.Staff.GetLevel(oPC.CurrentSecurityCode) < 2 Then
        MsgBox "You do not have security to alter rounding rules.", vbExclamation, "Denied"
        Exit Sub
    End If
    Set frm = New frmRR
    frm.Show 'vbModal
    Set frm = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRR_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSectionSales_Click()
    On Error GoTo errHandler
Dim oCU As z_Cashup
   ' oCU.PrintCashup

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSectionSales_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCopyDoc_Click()
    On Error GoTo errHandler
Dim frm As New frmBrowseCustomers2
    frm.Show vbModal
    Me.ActiveForm.CopyThisDoc frm.CustomerID
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCopyDoc_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSalesPatterns_Click()
    On Error GoTo errHandler
    frmBrowseProd.ShowSalesPatterns
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSalesPatterns_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSalesPatterns2_Click()
    On Error Resume Next
    If Me.ActiveForm Is frmREORDER_SAL Then
        frmREORDER_SAL.ShowSalesPatterns
    Else
        frmREORDER_CUST.ShowSalesPatterns
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSalesPatterns2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSaveColumnWidths_Click()
    On Error Resume Next
    Me.ActiveForm.mnuSaveLayout
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSaveColumnWidths_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub mnuSPA_Click()
    On Error GoTo errHandler
Dim frm As New frmSalesSummary
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSPA_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSTAT1_Click()
    On Error GoTo errHandler
MsgBox "This report is now available from the Papyrus II Reports application", vbOKOnly + vbInformation, "Information"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSTAT1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSpecialRequest_Click()
Dim lngStaffID As Long
    On Error Resume Next

    If SecurityControl(enSECURITY_ISSUPERVISOR, , "Enter your signature", "You need to be a supervisor to add to special orders", , , lngStaffID) Then
        Me.ActiveForm.mnuAddToSpecialOrder lngStaffID
    End If

End Sub

Private Sub mnuStatusChange_Click()
    On Error GoTo errHandler
Dim frm As New frmUpdateSupplierStatus
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuStatusChange_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSub_Click()
    NewPO "NS"
End Sub

Private Sub mnuSubstitute_Click()
    On Error Resume Next
    If Me.ActiveForm.Name = "frmInvoicePreview" Then
        Me.ActiveForm.InsertSubstitutes
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSubstitute_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuILine_COL_Click()
    On Error Resume Next
    Me.ActiveForm.ViewCOL
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuILine_COL_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuCopyLines_Click()
    On Error Resume Next
    Me.ActiveForm.mnuCopyLines
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCopyLines_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPasteLines_Click()
    On Error Resume Next
    Me.ActiveForm.mnuPastelines
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPasteLines_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPasteLinestoNewOrder_Click()
    On Error GoTo errHandler
    PastelinestoNEWOrder
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPasteLinestoNewOrder_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuCopytoNewPO_Click()
    On Error GoTo errHandler
    PastelinestoNEWPOOrder
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCopytoNewPO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuPasteLinestoNewPOOrder_Click()
    On Error GoTo errHandler
    PastelinestoNEWPOOrder
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPasteLinestoNewPOOrder_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPasteLinestoNewInvoice_Click()
    On Error GoTo errHandler
    PastelinestoNEWInvoice False, True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPasteLinestoNewInvoice_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPastelinestoNewCounterInvoice_Click()
    PastelinestoNEWInvoice False, False

End Sub
Private Sub mnuPastelinestoNEWPFInvoice_Click()
    On Error GoTo errHandler
    PastelinestoNEWInvoice True, False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPastelinestoNEWPFInvoice_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPasteLinestoNewQuotation_Click()
    On Error GoTo errHandler
    PastelinestoNEWQuotation
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPasteLinestoNewQuotation_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPastelinestoNEWAppro_Click()
    On Error GoTo errHandler
    PastelinestoNEWAppro
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPastelinestoNEWAppro_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuHeader_Click()
    On Error Resume Next
    Me.ActiveForm.mnuHeader
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuHeader_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuBrowseTActions_Click()
    On Error Resume Next
    Me.ActiveForm.BrowseTActions
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuBrowseTActions_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSupplierClaims_Click()
Dim frm As New frmBrowseSupplierClaims

    frm.Show
    
End Sub

'Private Sub mnuSwaptoTest_Click()
'    On Error GoTo errHandler
'Dim f As Form
'
'    Screen.MousePointer = vbHourglass
'    For Each f In Forms
'        If Not f Is Forms(0) Then Unload f
'    Next
'
'    oPC.SwapConnectionToDatabase
'    PaintFirstScreen
''    If oPC.DatabaseName = "PBKS_TEST" Then
''        Me.mnuSwaptoTest.Caption = "Swap to working on LIVE database"
''        Me.mnuNewTestFromLive.Enabled = False
''        Me.mnuManDBCopies.Enabled = False
''    Else
''        Me.mnuSwaptoTest.Caption = "Swap to working on TEST database"
''        Me.mnuNewTestFromLive.Enabled = True
''        Me.mnuManDBCopies.Enabled = True
''    End If
'    Screen.MousePointer = vbDefault
'
'    MsgBox "You are now connected to the " & IIf(oPC.DatabaseName = "PBKS_TEST", "TEST", "LIVE") & " database", vbOKOnly, "Status"
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.mnuSwaptoTest_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub mnuTextBites_Click()
    On Error Resume Next
    ShowTextBites
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTextBites_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPayInvoice_Click()
    On Error Resume Next
    Me.ActiveForm.mnuPayInvoice
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPayInvoice_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuTouchRecord_Click()
    On Error Resume Next
    Me.ActiveForm.mnuTouchRecord
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTouchRecord_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuAlertHistory_Click()
    On Error Resume Next
    Me.ActiveForm.mnuAlertHistory
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAlertHistory_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuAlert_Click()
    On Error Resume Next
    Me.ActiveForm.mnuAlert
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAlert_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuTransferIn_Click()
    On Error GoTo errHandler
    PastelinestoNEWTransfer ("IN")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTransferIn_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuTransferOut_Click()
    On Error GoTo errHandler
    PastelinestoNEWTransfer ("OUT")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTransferOut_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub mnuUpdatePOS_Cust_Click()
    On Error Resume Next
    Me.ActiveForm.mnuTouchRecord
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuUpdatePOS_Cust_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuTrans_Click()
    On Error GoTo errHandler
    BrowseTrans
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTrans_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuVoid_Click()
    On Error Resume Next
    Me.ActiveForm.mnuVoid
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuVoid_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub mnuMemo_Click()
'    Me.ActiveForm.mnuMemo
'End Sub
Private Sub mnuWants_Click()
    On Error GoTo errHandler
'Dim frm As New frmEntire
'    frm.Caption = "Create list of all books"
'    frm.txtDescription = "All books on database"
'    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuWants_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub mnuWeb_Click()
    On Error GoTo errHandler
Dim str As String
Dim str2 As String
    If oPC.InternetDialup = True Then Exit Sub
    Screen.MousePointer = vbHourglass
            str = "http://www.papyrussoftware.co.za"
            OpenBrowser str
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuWeb_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub SB1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
    If Shift = 1 Then
        SB1.Panels(2).text = oPC.NewQuotation
        SB1.Panels(2).ToolTipText = SB1.Panels(2).text
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.SB1_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), EA_NORERAISE
    HandleError
End Sub




Private Sub TBHead_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo errHandler
    Select Case UCase(Button.Key)
    Case "BINV"
        BrowseInvoices False
    Case "BAPP"
        BrowseApps
    Case "BAPPR"
        BrowseAPPRs
    Case "BDEL"
        BrowseDELS
    Case "BPO"
        BrowsePOs
    Case "BCO"
        BrowseOrders
    Case "BCN"
        BrowseCN
    Case "BGEN"
        BrowseGenStock
    Case "BBKS"
'190           If oPC.AllowAntiquarionSearch = 3 Then
'200               BrowseAntBooks
'210           Else
            BrowseBooks
      '  End If
    Case "BTR"
        BrowseTrans
    Case "BCUST"
        BrowseCustomers
    Case "BSUPP"
        BrowseSupplier
    Case "BCS"
        BrowseCS
    Case "NINV"
        NewInvoice False, False
    Case "NCUST"
        NewCustomer enPrivate
    Case "NGS"
        NewGenStock
    Case "NBK"
        NewBook
    Case "NSUPP"
        NewSupplier
    Case "NAPP"
        NewAPP
    Case "NAPPR"
        NewAPPR
    Case "NDEL"
        NewDEL
    Case "NPO"
        NewPO "R"
    Case "NCO"
        NewCO
    Case "NCN"
        NewCN
    Case "NTR"
       ' NewTRANS
    Case "NCN"
        NewCN
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.TBHead_ButtonClick(Button)", Button, EA_NORERAISE
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
    
'    Dim arCommandLine() As String

    ' Get the MDIClient area's device context
    aDC = GetDC(hWnd)
    ' Get the MDIClient dimensions
    GetWindowRect hWnd, rcClient
    ' shift the origin to 0,0
    rcClient.Right = rcClient.Right - rcClient.Left
    rcClient.Bottom = rcClient.Bottom - rcClient.TOP
    rcClient.TOP = 0
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
        If UBound(arCommandLine) > 0 Then
            If arCommandLine(1) <> "N" Then
                abrush = CreateSolidBrush(vbRed)
            Else
                abrush = CreateSolidBrush(RGB(25, 38, 85))
            End If
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
    ReleaseDC hWnd, aDC
'Errh:
'    MsgBox Error
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

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.Toolbar2_ButtonClick(Button)", Button, EA_NORERAISE
    HandleError
End Sub
Private Sub mnuAlloc_Click()
    On Error GoTo errHandler
Dim frmS As frmOrderFUlfil_Selection
Dim frmAlloc As frmCOLAllocation
Dim cCOLALLOC As chex_COLAllocation
Dim bComplete As Boolean
Dim strWSName As String

    strWSName = oPC.NameOfPC

    Set cCOLALLOC = Nothing
    Set cCOLALLOC = New chex_COLAllocation
    
    Set frmS = New frmOrderFUlfil_Selection
    frmS.Show vbModal
    bComplete = frmS.CompleteOnly
    If Not frmS.CancelledYN Then
        WaitMsg "Fetching . . . ", True
        If frmS.LoadLastSet Then
        ElseIf frmS.oPCode > 0 Then
            cCOLALLOC.GenerateCOLAllocationset , frmS.oPCode, , , , , bComplete
        ElseIf frmS.CustID > 0 Then
            cCOLALLOC.GenerateCOLAllocationset , , frmS.CustID, , , , bComplete
        ElseIf frmS.SupplierID > 0 Then
            cCOLALLOC.GenerateCOLAllocationset , , , frmS.SupplierID, , , bComplete
        ElseIf frmS.CustFrom > "" And frmS.CustTo > "" Then
            cCOLALLOC.GenerateCOLAllocationset , , , , frmS.CustFrom, frmS.CustTo, bComplete
        End If
    
        cCOLALLOC.Load , , strWSName
        WaitMsg "", False
        If cCOLALLOC.Count > 0 Then
            Set frmAlloc = New frmCOLAllocation
            frmAlloc.component cCOLALLOC, "NORMAL", False, IIf(oPC.IncludeSupplierFeatures, "G", "I")
            frmAlloc.Show
        Else
            If frmS.LoadLastSet Then
                MsgBox "There are no orders to load.", vbInformation + vbOKOnly, "Status"
            Else
                MsgBox "There are no orders matching your criteria.", vbInformation + vbOKOnly, "Status"
            End If
        End If
    End If
    Unload frmS


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAlloc_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuOrderFulfilmentAppros_Click()
Dim frmS As frmOrderFUlfil_Selection
Dim frmAlloc As frmCOLAllocation
Dim cCOLALLOC As chex_COLAllocation
Dim bComplete As Boolean
Dim strWSName As String

    strWSName = oPC.NameOfPC

    Set cCOLALLOC = Nothing
    Set cCOLALLOC = New chex_COLAllocation
    
    Set frmS = New frmOrderFUlfil_Selection
    frmS.Show vbModal
    bComplete = frmS.CompleteOnly
    If Not frmS.CancelledYN Then
        WaitMsg "Fetching . . . ", True
        If frmS.LoadLastSet Then
        ElseIf frmS.oPCode > 0 Then
            cCOLALLOC.GenerateCOLAllocationset , frmS.oPCode, , , , , bComplete
        ElseIf frmS.CustID > 0 Then
            cCOLALLOC.GenerateCOLAllocationset , , frmS.CustID, , , , bComplete
        ElseIf frmS.SupplierID > 0 Then
            cCOLALLOC.GenerateCOLAllocationset , , , frmS.SupplierID, , , bComplete
        ElseIf frmS.CustFrom > "" And frmS.CustTo > "" Then
            cCOLALLOC.GenerateCOLAllocationset , , , , frmS.CustFrom, frmS.CustTo, bComplete
        End If
    
        cCOLALLOC.Load , , strWSName
        WaitMsg "", False
        If cCOLALLOC.Count > 0 Then
            Set frmAlloc = New frmCOLAllocation
            frmAlloc.component cCOLALLOC, "NORMAL", False, "A"
            frmAlloc.Show
        Else
            If frmS.LoadLastSet Then
                MsgBox "There are no orders to load.", vbInformation + vbOKOnly, "Status"
            Else
                MsgBox "There are no orders matching your criteria.", vbInformation + vbOKOnly, "Status"
            End If
        End If
    End If
    Unload frmS


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuOrderFulfilmentAppros_Click", , EA_NORERAISE
    HandleError

End Sub

Private Sub mnuClearTempList_Click()
    On Error GoTo errHandler
    frmBrowseProd.StartNewList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuClearTempList_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuAdd_Click()
    On Error GoTo errHandler
    frmBrowseProd.AddToTempList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuAdd_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuCOPlace_Click()
    On Error GoTo errHandler
    frmBrowseProd.PlaceCO
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCOPlace_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuReorderFromBrowse_Click()
    On Error GoTo errHandler
    frmBrowseProd.mnuLoadReorderSlate
    Set frmREORDER_SAL = New frmREORDER_CO
    frmREORDER_SAL.component "BROWSED"
    frmREORDER_SAL.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuReorderFromBrowse_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuRemoveFromReorderList_Click()
    On Error Resume Next
    If Me.ActiveForm Is frmREORDER_SAL Then
        frmREORDER_SAL.RemoveFromList
    Else
        frmREORDER_CUST.RemoveFromList
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRemoveFromReorderList_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuRejected_Click()
    On Error Resume Next
    Screen.ActiveForm.ManageReturnRejection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuRejected_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuExportCat_Click()
    On Error Resume Next
    Screen.ActiveForm.ExportInCatalogueFormat
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuExportCat_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuSetDeal_Click()
    On Error Resume Next
    If Me.ActiveForm Is frmREORDER_SAL Then
        If Not frmREORDER_SAL Is Nothing Then frmREORDER_SAL.SetDeal
    Else
        If Me.ActiveForm Is frmREORDER_CUST Then
            If Not frmREORDER_CUST Is Nothing Then frmREORDER_CUST.SetDeal
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSetDeal_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuCustomerCollects_Click()
    On Error Resume Next
    Screen.ActiveForm.CustomerCollects
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuCustomerCollects_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPutBack_Click()
    On Error Resume Next
    Screen.ActiveForm.ReturnToStock
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPutBack_Click", , EA_NORERAISE
    HandleError
End Sub
'''''''''''''''''''''''''
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
Private Sub NewTransfer(pDir As String)
    On Error GoTo errHandler
Dim frm1 As frmTFR2
Dim lngTPID As Long
Dim frm As frmBrowseStores
Dim bCancel As Boolean
Dim strStorename As String
Dim errRepeat As Integer

    errSysHandlerSet

    Set frm = New frmBrowseStores
    frm.Show vbModal
    lngTPID = frm.StoreID
    strStorename = frm.StoreName
    Unload frm
    If lngTPID = 0 Then Exit Sub
    Set frm1 = New frmTFR2
    frm1.component pDir, bCancel, lngTPID, , strStorename
    If bCancel Then
  Unload frm1
    Else
  frm1.Show
    End If

    Exit Sub
errHandler:
    ErrPreserve
    Set frm = Nothing
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmMain: NewTransfer, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmMain: NewTransfer after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.NewTransfer(pDir)", pDir
End Sub

Private Sub mnuTranOut_Click()
    On Error GoTo errHandler
    
    NewTransfer "OUT"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTranOut_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuTFRIn_Click()
    On Error GoTo errHandler
    
    NewTransfer "IN"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuTFRIn_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuShowPOLHist_Click()
    On Error Resume Next
    Screen.ActiveForm.ShowPreviousOLVersions
    If Err Then
        MsgBox "Select the order row before trying to see previous versions.", vbInformation, "Can't do this"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuShowPOLHist_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuShowCOLHist_Click()
    On Error Resume Next
    Screen.ActiveForm.ShowPreviousOLVersions
    If Err Then
        MsgBox "Select the order row before trying to see previous versions.", vbInformation, "Can't do this"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuShowCOLHist_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPreDelAdv_Click()
    On Error Resume Next
    Screen.ActiveForm.mnuPreDelAdv
    If Err Then
        MsgBox "Select the order row before trying to generate messages.", vbInformation, "Can't do this"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPreDelAdv_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuSetSection_Click()
    On Error Resume Next
    If SecurityControl(enSECURITY_EDITPRODUCTTYPES_AUTH, , "Enter your signature", "You need permission to work with product types/sections") Then
        Me.ActiveForm.mnuSetSection
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSetSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub mnuSetPT_Click()
    On Error Resume Next
    If SecurityControl(enSECURITY_EDITPRODUCTTYPES_AUTH, , "Enter your signature", "You need permission to work with product types/sections") Then
        Me.ActiveForm.mnuSetPT
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuSetPT_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuPrintSelected_Click()
    Me.ActiveForm.PrintSelected
End Sub
Private Sub mnuPrintSelectedDeliveries_Click()
    Me.ActiveForm.PrintSelectedDeliveries
End Sub
Private Sub mnuEmail_Click()
    On Error Resume Next
    Me.ActiveForm.mnuEmail
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuEmail_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuOutlook_Click()
    On Error Resume Next
    Me.ActiveForm.mnuOutlook
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuOutlook_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub mnuEDI_Click()
    On Error Resume Next
    Me.ActiveForm.mnuEDI
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuEDI_Click", , EA_NORERAISE
    HandleError
End Sub


Public Sub PastelinestoNEWInvoice(IsProforma As Boolean, IsPreDelivery As Boolean)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngQUID As Long
Dim oSM As New z_StockManager
Dim oInv As New a_Invoice
Dim lngTPID As Long
Dim lngOrderID As Long
Dim frmInvPre As frmInvoicePreview
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim Qty As Long


    If oPC.QtyLinesinClipboard <= 0 Then
        If MsgBox("There are no document lines in the Papyrus clipboard. If you choose to continue, you will create a blank invoice." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("You have chosen to create an invoice using the " & CStr(oPC.QtyLinesinClipboard) & " lines in the Papyrus clipboard." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub

    lngOrderID = oSM.CreateNewInvoice(lngTPID, IsProforma, IsPreDelivery)
    
    Set oInv = Nothing
    Set oInv = New a_Invoice
    
    oInv.Load lngOrderID, False
'    oInv.BeginEdit
'    If (IsProforma = True) Then oInv.SetProforma
'    oInv.ApplyEdit
    Set rs = oPC.LinesClipboard
    If rs.State <> 0 Then
        If rs.BOF And rs.eof Then Exit Sub
        rs.MoveFirst
        Do While Not rs.eof
            If FNN(rs.Fields("QTYFIRM")) > 0 Then
                Qty = FNN(rs.Fields("QTYFIRM"))
            Else
                Qty = FNN(rs.Fields("QTY"))
            End If
        
            oInv.PasteLine FNS(rs.Fields("PID")), Qty, FNN(rs.Fields("QTYSS")), FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), FNDBL(rs.Fields("VATRATE")), _
                        FNS(rs.Fields("REF")), FNS(rs.Fields("EXTRACHARGEPID")), FNN(rs.Fields("EXTRACHARGEVALUE")), _
                    FNN(rs.Fields("FCPRICE")), FNDBL(rs.Fields("FCFACTOR")), FNN(rs.Fields("FCID"))
            
            rs.MoveNext
        Loop
    End If
    Set oInv = Nothing
    Set oInv = New a_Invoice
    
    oInv.Load lngOrderID, False
    
    
    Set frmInvPre = New frmInvoicePreview
    frmInvPre.ComponentObject oInv
    
    frmInvPre.Show
    
EXIT_Handler:
    Me.MousePointer = vbDefault
    Exit Sub
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.PastelinestoNEWInvoice"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.PastelinestoNEWInvoice(Proforma)", Proforma
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PastelinestoNEWInvoice(IsProforma,IsPreDelivery)", Array(IsProforma, _
         IsPreDelivery)
End Sub
Public Sub PastelinestoNEWOrder()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngQUID As Long
Dim oSM As New z_StockManager
Dim oCO As New a_CO
Dim lngTPID As Long
Dim lngOrderID As Long
Dim frmCOPre As frmCOPreview
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim Qty As Long
    If oPC.QtyLinesinClipboard <= 0 Then
        If MsgBox("There are no document lines in the Papyrus clipboard. If you choose to continue, you will create a blank order." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("You have chosen to create an order using the " & CStr(oPC.QtyLinesinClipboard) & " lines in the Papyrus clipboard." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If

    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    lngOrderID = oSM.CreateNewOrder(lngTPID)
    Set oCO = Nothing
    Set oCO = New a_CO
    oCO.Load lngOrderID, False
    Set rs = oPC.LinesClipboard
    If rs.State <> 0 Then
        If rs.BOF And rs.eof Then Exit Sub
        rs.MoveFirst
        Do While Not rs.eof
            If FNN(rs.Fields("QTYFIRM")) > 0 Then
                Qty = FNN(rs.Fields("QTYFIRM"))
            Else
                Qty = FNN(rs.Fields("QTY"))
            End If
            oCO.PasteLine FNS(rs.Fields("PID")), Qty, FNN(rs.Fields("QTYSS")), FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), _
                    FNDBL(rs.Fields("VATRATE")), FNS(rs.Fields("REF")), FNS(rs.Fields("EXTRACHARGEPID")), FNN(rs.Fields("EXTRACHARGEVALUE")), DateAdd("ww", 3, Date), _
                    FNN(rs.Fields("FCPRICE")), FNDBL(rs.Fields("FCFACTOR")), FNN(rs.Fields("FCID"))
            rs.MoveNext
        Loop
    End If
    Set oCO = Nothing
    Set oCO = New a_CO
    oCO.Load lngOrderID, False
    Set frmCOPre = New frmCOPreview
    frmCOPre.ComponentObject oCO
    frmCOPre.Show
    
EXIT_Handler:
    Me.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.PastelinestoNEWOrder"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PastelinestoNEWOrder"
End Sub
Public Sub PastelinestoNEWTransfer(pInOut As String)
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngQUID As Long
Dim oSM As New z_StockManager
Dim oTFR As New a_TF
Dim lngTPID As Long
Dim lngTFRID As Long
Dim frmTFRPre As frmTFPreview
Dim frm As frmBrowseStores
Dim bCancel As Boolean
Dim Qty As Long
Dim strStorename As String

    If oPC.QtyLinesinClipboard <= 0 Then
        If MsgBox("There are no document lines in the Papyrus clipboard. If you choose to continue, you will create a blank transfer." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("You have chosen to create a transfer using the " & CStr(oPC.QtyLinesinClipboard) & " lines in the Papyrus clipboard." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If

    Set frm = New frmBrowseStores
    frm.Show vbModal
    lngTPID = frm.StoreID
    strStorename = frm.StoreName
    Unload frm
    If lngTPID = 0 Then Exit Sub

    lngTFRID = oSM.CreateNewTransfer(lngTPID, pInOut)
    Set oTFR = Nothing
    Set oTFR = New a_TF
    oTFR.Load lngTFRID
    Set rs = oPC.LinesClipboard
    If rs.State <> 0 Then
        If rs.BOF And rs.eof Then Exit Sub
        rs.MoveFirst
        Do While Not rs.eof
            oTFR.PasteLine FNS(rs.Fields("PID")), FNN(rs.Fields("QTYFIRM")) + FNN(rs.Fields("QTYSS")), FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), _
                    FNDBL(rs.Fields("VATRATE"))
            rs.MoveNext
        Loop
    End If
    Set oTFR = Nothing
    Set oTFR = New a_TF
    oTFR.Load lngTFRID
    Set frmTFRPre = New frmTFPreview
    frmTFRPre.ComponentObject oTFR
    frmTFRPre.Show
    
EXIT_Handler:
    Me.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.PastelinestoNEWTransferIn"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PastelinestoNEWTransfer(pInOut)", pInOut
End Sub

Public Sub PastelinestoNEWPOOrder()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngQUID As Long
Dim oSM As New z_StockManager
Dim oPO As New a_PO
Dim lngTPID As Long
Dim lngOrderID As Long
Dim frmPOPre As frmPOPreview
Dim frm As frmBrowseSUppliers2
Dim bCancel As Boolean
Dim Qty As Long
    If oPC.QtyLinesinClipboard <= 0 Then
        If MsgBox("There are no document lines in the Papyrus clipboard. If you choose to continue, you will create a blank order." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("You have chosen to create an order using the " & CStr(oPC.QtyLinesinClipboard) & " lines in the Papyrus clipboard." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If

    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    lngTPID = frm.SupplierID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    lngOrderID = oSM.CreateNewPOOrder(lngTPID)
    Set oPO = Nothing
    Set oPO = New a_PO
    oPO.Load lngOrderID, False
    Set rs = oPC.LinesClipboard
    If rs.State <> 0 Then
        If rs.BOF And rs.eof Then Exit Sub
        rs.MoveFirst
        Do While Not rs.eof
            If FNN(rs.Fields("QTYFIRM")) > 0 Then
                Qty = FNN(rs.Fields("QTYFIRM"))
            Else
                Qty = FNN(rs.Fields("QTY"))
            End If
            oPO.PasteLine FNS(rs.Fields("PID")), Qty, FNN(rs.Fields("QTYSS")), FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), _
                    FNDBL(rs.Fields("VATRATE")), FNS(rs.Fields("REF")), FND(rs.Fields("ETA"))
            rs.MoveNext
        Loop
    End If
    Set oPO = Nothing
    Set oPO = New a_PO
    oPO.Load lngOrderID, False
    Set frmPOPre = New frmPOPreview
    frmPOPre.component oPO.TRID
    frmPOPre.Show
    
EXIT_Handler:
    Me.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.PastelinestoNEWOrder"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PastelinestoNEWPOOrder"
End Sub

Public Sub PastelinestoNEWQuotation()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngQUID As Long
Dim oSM As New z_StockManager
Dim oQU As New a_QU
Dim lngTPID As Long
Dim lngID As Long
Dim frmQUPre As frmQuotationPreview
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim Qty As Long
    If oPC.QtyLinesinClipboard <= 0 Then
        If MsgBox("There are no document lines in the Papyrus clipboard. If you choose to continue, you will create a blank quotation." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("You have chosen to create an quotation using the " & CStr(oPC.QtyLinesinClipboard) & " lines in the Papyrus clipboard." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If

    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    lngID = oSM.CreateNewQuotation(lngTPID)
    Set oQU = Nothing
    Set oQU = New a_QU
    oQU.Load lngID, False
    Set rs = oPC.LinesClipboard
    If rs.State <> 0 Then
        If rs.BOF And rs.eof Then Exit Sub
        rs.MoveFirst
        Do While Not rs.eof
            If FNN(rs.Fields("QTYFIRM")) > 0 Then
                Qty = FNN(rs.Fields("QTYFIRM"))
            Else
                Qty = FNN(rs.Fields("QTY"))
            End If
            oQU.PasteLine FNS(rs.Fields("PID")), Qty, FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), FNDBL(rs.Fields("VATRATE")), FNS(rs.Fields("REF")), _
                        FNS(rs.Fields("EXTRACHARGEPID")), FNN(rs.Fields("EXTRACHARGEVALUE")), _
                        FNN(rs.Fields("FCPRICE")), FNDBL(rs.Fields("FCFACTOR")), FNN(rs.Fields("FCID"))
            rs.MoveNext
        Loop
    End If
    Set oQU = Nothing
    Set oQU = New a_QU
    oQU.Load lngID, False
    Set frmQUPre = New frmQuotationPreview
    frmQUPre.ComponentObject oQU
    frmQUPre.Show
    
EXIT_Handler:
    Me.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.PastelinestoNEWQuotation"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PastelinestoNEWQuotation"
End Sub

Public Sub PastelinestoNEWAppro()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset
Dim lngQUID As Long
Dim oSM As New z_StockManager
Dim oAPP As New a_APP
Dim lngTPID As Long
Dim lngID As Long
Dim frmAppPre As frmAPPPreview
Dim frm As frmBrowseCustomers2
Dim bCancel As Boolean
Dim Qty As Long
    If oPC.QtyLinesinClipboard <= 0 Then
        If MsgBox("There are no document lines in the Papyrus clipboard. If you choose to continue, you will create a blank appro." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    If MsgBox("You have chosen to create an appro using the " & CStr(oPC.QtyLinesinClipboard) & " lines in the Papyrus clipboard." & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If

    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    lngTPID = frm.CustomerID
    Unload frm
    If lngTPID = 0 Then Exit Sub
    lngID = oSM.CreateNewAppro(lngTPID)
    Set oAPP = Nothing
    Set oAPP = New a_APP
    oAPP.Load lngID, False
    Set rs = oPC.LinesClipboard
    If rs.State <> 0 Then
        If rs.BOF And rs.eof Then Exit Sub
        rs.MoveFirst
        Do While Not rs.eof
            If FNN(rs.Fields("QTYFIRM")) > 0 Then
                Qty = FNN(rs.Fields("QTYFIRM"))
            Else
                Qty = FNN(rs.Fields("QTY"))
            End If
            oAPP.PasteLine FNS(rs.Fields("PID")), Qty, FNN(rs.Fields("PRICE")), FNDBL(rs.Fields("DISCOUNTRATE")), FNDBL(rs.Fields("VATRATE")), FNS(rs.Fields("REF"))
            rs.MoveNext
        Loop
    End If
    Set oAPP = Nothing
    Set oAPP = New a_APP
    oAPP.Load lngID, False
    Set frmAppPre = New frmAPPPreview
    frmAppPre.ComponentObject oAPP
    frmAppPre.Show
    
EXIT_Handler:
    Me.MousePointer = vbDefault
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmMain.PastelinestoNEWQuotation"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.PastelinestoNEWAppro"
End Sub

Private Sub mnuPrepareDetailList_Click()
    On Error Resume Next
    Me.ActiveForm.PrepareDetailList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmMain.mnuPrepareDetailList_Click", , EA_NORERAISE
    HandleError
End Sub
