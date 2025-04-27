VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmConfiguration 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Configuration"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   10680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   10680
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   180
      Picture         =   "frmConfiguration.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   6300
      Width           =   1000
   End
   Begin VB.CommandButton cmdTestErrorHandling_DeveloperOnly 
      Caption         =   "Command2"
      Height          =   360
      Left            =   225
      TabIndex        =   69
      Top             =   7950
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8310
      Picture         =   "frmConfiguration.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6300
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9330
      Picture         =   "frmConfiguration.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6300
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6120
      Left            =   195
      TabIndex        =   3
      Top             =   120
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   10795
      _Version        =   393216
      Tabs            =   12
      Tab             =   6
      TabsPerRow      =   6
      TabHeight       =   670
      BackColor       =   13489106
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Main"
      TabPicture(0)   =   "frmConfiguration.frx":0A9E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblCSCustomer"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label27"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "chkAllowCopy"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkSignTransactions"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtOfferSignature"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtVATRate"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPrefix"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkCC"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cboLocalCountry"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtLookupSeq"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "chkCaptureDecimal"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chkAntiquarian"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdLocateCS"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkDiscountVATDefault"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkSections"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chkSupportsWORD"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkNonBookYN"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtCOLAllocStyle"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkReorderStyle"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "chkEnforceCOLRefs"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkAggregatePOs"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "chkLoyalty"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdCopyBFFiles"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtEDINumber"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtMU"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Companies"
      TabPicture(1)   =   "frmConfiguration.frx":0ABA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvwCompanies"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdAddComp"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdEditComp"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdDefault"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdRemove"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Currencies"
      TabPicture(2)   =   "frmConfiguration.frx":0AD6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwCurrencies"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdRemCurr"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdDefaultCurr"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdEditCurr"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdAddCurr"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdLocal"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Stores"
      TabPicture(3)   =   "frmConfiguration.frx":0AF2
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvwStores"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "cmdSetDefaultStore"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdStoreEdit"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdAddStore"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdBillToStore"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Misc."
      TabPicture(4)   =   "frmConfiguration.frx":0B0E
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Frame2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Frame3"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame5"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      TabCaption(5)   =   "Documents"
      TabPicture(5)   =   "frmConfiguration.frx":0B2A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label9"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label41"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Grid1"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "DDPrinters"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "cboWorkstations"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "cmdLoadDefaults"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "cmdDeletePrinter"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "lvwPrinters"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).ControlCount=   8
      TabCaption(6)   =   "Staff"
      TabPicture(6)   =   "frmConfiguration.frx":0B46
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "lvwStaff"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "cmdAddstaff"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "cmdEdiTsTAFF"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "cmdRemoveStaff"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).ControlCount=   4
      TabCaption(7)   =   "Fixed text"
      TabPicture(7)   =   "frmConfiguration.frx":0B62
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "txtStatementMessage"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "txt"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "txtInvText"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "txtQUText"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "txtCOText"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "txtPOText"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "Label35"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "Label34"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "Label16"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "Label21"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "Label31"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).Control(11)=   "Label32"
      Tab(7).Control(11).Enabled=   0   'False
      Tab(7).Control(12)=   "Label33"
      Tab(7).Control(12).Enabled=   0   'False
      Tab(7).ControlCount=   13
      TabCaption(8)   =   "Imp/Exp"
      TabPicture(8)   =   "frmConfiguration.frx":0B7E
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "txtGLReference"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "txtCreditorsContra"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).Control(2)=   "txtDebtorsContra"
      Tab(8).Control(2).Enabled=   0   'False
      Tab(8).Control(3)=   "Frame4"
      Tab(8).Control(3).Enabled=   0   'False
      Tab(8).Control(4)=   "Label22"
      Tab(8).Control(4).Enabled=   0   'False
      Tab(8).Control(5)=   "Label20"
      Tab(8).Control(5).Enabled=   0   'False
      Tab(8).Control(6)=   "Label17"
      Tab(8).Control(6).Enabled=   0   'False
      Tab(8).ControlCount=   7
      TabCaption(9)   =   "Defaults"
      TabPicture(9)   =   "frmConfiguration.frx":0B9A
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "cboDefCategory"
      Tab(9).Control(0).Enabled=   0   'False
      Tab(9).Control(1)=   "cboDefPT"
      Tab(9).Control(1).Enabled=   0   'False
      Tab(9).Control(2)=   "cboDefCustomerType"
      Tab(9).Control(2).Enabled=   0   'False
      Tab(9).Control(3)=   "cboIGLaunch"
      Tab(9).Control(3).Enabled=   0   'False
      Tab(9).Control(4)=   "cboIGPromotion"
      Tab(9).Control(4).Enabled=   0   'False
      Tab(9).Control(5)=   "cboIGSale"
      Tab(9).Control(5).Enabled=   0   'False
      Tab(9).Control(6)=   "cboIGLunch"
      Tab(9).Control(6).Enabled=   0   'False
      Tab(9).Control(7)=   "Label30"
      Tab(9).Control(7).Enabled=   0   'False
      Tab(9).Control(8)=   "Label29"
      Tab(9).Control(8).Enabled=   0   'False
      Tab(9).Control(9)=   "Label28"
      Tab(9).Control(9).Enabled=   0   'False
      Tab(9).Control(10)=   "Label11"
      Tab(9).Control(10).Enabled=   0   'False
      Tab(9).Control(11)=   "Label12"
      Tab(9).Control(11).Enabled=   0   'False
      Tab(9).Control(12)=   "Label13"
      Tab(9).Control(12).Enabled=   0   'False
      Tab(9).Control(13)=   "Label15"
      Tab(9).Control(13).Enabled=   0   'False
      Tab(9).ControlCount=   14
      TabCaption(10)  =   "Central"
      TabPicture(10)  =   "frmConfiguration.frx":0BB6
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame1"
      Tab(10).Control(0).Enabled=   0   'False
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Email message"
      TabPicture(11)  =   "frmConfiguration.frx":0BD2
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "txtEmailQuotationMsg"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).Control(1)=   "txtEmailApproMsg"
      Tab(11).Control(1).Enabled=   0   'False
      Tab(11).Control(2)=   "txtEMailPOMsg"
      Tab(11).Control(2).Enabled=   0   'False
      Tab(11).Control(3)=   "txtEmailInvoiceMsg"
      Tab(11).Control(3).Enabled=   0   'False
      Tab(11).Control(4)=   "Label40"
      Tab(11).Control(4).Enabled=   0   'False
      Tab(11).Control(5)=   "Label39"
      Tab(11).Control(5).Enabled=   0   'False
      Tab(11).Control(6)=   "Label38"
      Tab(11).Control(6).Enabled=   0   'False
      Tab(11).Control(7)=   "Label37"
      Tab(11).Control(7).Enabled=   0   'False
      Tab(11).Control(8)=   "Label36"
      Tab(11).Control(8).Enabled=   0   'False
      Tab(11).ControlCount=   9
      Begin MSComctlLib.ListView lvwPrinters 
         Height          =   3720
         Left            =   -67905
         TabIndex        =   139
         Top             =   1530
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   6562
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Printer"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.CommandButton cmdBillToStore 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Bill to store"
         Height          =   420
         Left            =   -68205
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   1980
         Width           =   1500
      End
      Begin VB.TextBox txtEmailQuotationMsg 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -70110
         MultiLine       =   -1  'True
         TabIndex        =   135
         Top             =   2595
         Width           =   4425
      End
      Begin VB.TextBox txtEmailApproMsg 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -70170
         MultiLine       =   -1  'True
         TabIndex        =   133
         Top             =   1215
         Width           =   4425
      End
      Begin VB.TextBox txtEMailPOMsg 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -74775
         MultiLine       =   -1  'True
         TabIndex        =   129
         Top             =   1215
         Width           =   4425
      End
      Begin VB.TextBox txtEmailInvoiceMsg 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -74730
         MultiLine       =   -1  'True
         TabIndex        =   128
         Top             =   2595
         Width           =   4425
      End
      Begin VB.CommandButton cmdDeletePrinter 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Delete selected printer"
         Height          =   420
         Left            =   -67905
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   5265
         Width           =   2715
      End
      Begin VB.TextBox txtStatementMessage 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -74820
         MultiLine       =   -1  'True
         TabIndex        =   120
         Top             =   3660
         Width           =   4425
      End
      Begin VB.TextBox txt 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -69990
         MultiLine       =   -1  'True
         TabIndex        =   119
         Top             =   3645
         Width           =   4425
      End
      Begin VB.TextBox txtInvText 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -69990
         MultiLine       =   -1  'True
         TabIndex        =   117
         Top             =   2415
         Width           =   4425
      End
      Begin VB.TextBox txtQUText 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -69975
         MultiLine       =   -1  'True
         TabIndex        =   116
         Top             =   1170
         Width           =   4425
      End
      Begin VB.TextBox txtCOText 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -74805
         MultiLine       =   -1  'True
         TabIndex        =   115
         Top             =   2415
         Width           =   4425
      End
      Begin VB.Frame Frame1 
         Caption         =   "Loyalty numbering"
         ForeColor       =   &H8000000D&
         Height          =   2370
         Left            =   -74520
         TabIndex        =   109
         Top             =   1155
         Width           =   4920
         Begin VB.TextBox txtEndLoy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3120
            TabIndex        =   112
            Top             =   1815
            Width           =   990
         End
         Begin VB.TextBox txtStartLoy 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3120
            TabIndex        =   110
            Top             =   1335
            Width           =   990
         End
         Begin VB.Label Label19 
            Caption         =   $"frmConfiguration.frx":0BEE
            ForeColor       =   &H8000000D&
            Height          =   870
            Left            =   540
            TabIndex        =   114
            Top             =   345
            Width           =   3645
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "Loyalty numbering end"
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   990
            TabIndex        =   113
            Top             =   1845
            Width           =   1965
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Loyalty numbering start"
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   990
            TabIndex        =   111
            Top             =   1365
            Width           =   1965
         End
      End
      Begin VB.ComboBox cboDefCategory 
         Height          =   315
         Left            =   -72330
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   2070
         Width           =   1830
      End
      Begin VB.ComboBox cboDefPT 
         Height          =   315
         Left            =   -72345
         TabIndex        =   105
         Top             =   1575
         Width           =   1830
      End
      Begin VB.ComboBox cboDefCustomerType 
         Height          =   315
         ItemData        =   "frmConfiguration.frx":0CB1
         Left            =   -72330
         List            =   "frmConfiguration.frx":0CB3
         TabIndex        =   103
         Top             =   1110
         Width           =   1830
      End
      Begin VB.ComboBox cboIGLaunch 
         Height          =   315
         Left            =   -67245
         TabIndex        =   98
         Top             =   990
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboIGPromotion 
         Height          =   315
         Left            =   -67245
         TabIndex        =   97
         Top             =   1455
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboIGSale 
         Height          =   315
         Left            =   -67245
         TabIndex        =   96
         Top             =   1935
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.ComboBox cboIGLunch 
         Height          =   315
         Left            =   -67260
         TabIndex        =   95
         Top             =   2400
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.CommandButton cmdRemoveStaff 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   5115
         Width           =   1095
      End
      Begin VB.TextBox txtMU 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -67860
         TabIndex        =   92
         Top             =   2325
         Width           =   915
      End
      Begin VB.Frame Frame5 
         Caption         =   "Period cutoffs"
         ForeColor       =   &H8000000D&
         Height          =   4725
         Left            =   -69660
         TabIndex        =   82
         Top             =   1065
         Visible         =   0   'False
         Width           =   3675
         Begin VB.CommandButton cmdRolldates 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Roll dates for month-end"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   930
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   4050
            Width           =   1935
         End
         Begin MSComCtl2.DTPicker dpCur 
            Height          =   375
            Left            =   480
            TabIndex        =   83
            Top             =   630
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   208207873
            CurrentDate     =   38946
         End
         Begin MSComCtl2.DTPicker dp30 
            Height          =   375
            Left            =   480
            TabIndex        =   85
            Top             =   1540
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   208207873
            CurrentDate     =   38946
         End
         Begin MSComCtl2.DTPicker dp60 
            Height          =   375
            Left            =   480
            TabIndex        =   87
            Top             =   2450
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   208207873
            CurrentDate     =   38946
         End
         Begin MSComCtl2.DTPicker dp90 
            Height          =   375
            Left            =   480
            TabIndex        =   89
            Top             =   3360
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   208207873
            CurrentDate     =   38946
         End
         Begin VB.Label Label26 
            Caption         =   "90 days period began at midnight on:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   480
            TabIndex        =   90
            ToolTipText     =   "End of period during which the update can commence"
            Top             =   3120
            Width           =   2820
         End
         Begin VB.Label Label25 
            Caption         =   "60 days period began at midnight on:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   480
            TabIndex        =   88
            ToolTipText     =   "End of period during which the update can commence"
            Top             =   2210
            Width           =   2940
         End
         Begin VB.Label Label24 
            Caption         =   "30 days period began at midnight on:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   480
            TabIndex        =   86
            ToolTipText     =   "End of period during which the update can commence"
            Top             =   1300
            Width           =   2760
         End
         Begin VB.Label Label23 
            Caption         =   "Current month began at midnight on:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   480
            TabIndex        =   84
            ToolTipText     =   "End of period during which the update can commence"
            Top             =   390
            Width           =   2760
         End
      End
      Begin VB.TextBox txtGLReference 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67290
         TabIndex        =   80
         Top             =   2265
         Width           =   1935
      End
      Begin VB.TextBox txtCreditorsContra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67290
         TabIndex        =   78
         Top             =   1785
         Width           =   1935
      End
      Begin VB.TextBox txtDebtorsContra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67290
         TabIndex        =   76
         Top             =   1305
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "3rd party application"
         ForeColor       =   &H8000000D&
         Height          =   1695
         Left            =   -74415
         TabIndex        =   72
         Top             =   1275
         Width           =   2265
         Begin VB.OptionButton optNone 
            Caption         =   "None"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   240
            TabIndex        =   75
            Top             =   1170
            Width           =   1665
         End
         Begin VB.OptionButton optPastel 
            Caption         =   "Pastel"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   240
            TabIndex        =   74
            Top             =   390
            Value           =   -1  'True
            Width           =   1665
         End
         Begin VB.OptionButton optAccpac 
            Caption         =   "Accpac"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   240
            TabIndex        =   73
            Top             =   780
            Width           =   1665
         End
      End
      Begin VB.TextBox txtPOText 
         ForeColor       =   &H8000000D&
         Height          =   885
         Left            =   -74790
         MultiLine       =   -1  'True
         TabIndex        =   70
         Top             =   1170
         Width           =   4425
      End
      Begin VB.TextBox txtEDINumber 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73035
         TabIndex        =   67
         Top             =   2280
         Width           =   2655
      End
      Begin VB.CommandButton cmdCopyBFFiles 
         BackColor       =   &H00C4BCA4&
         Cancel          =   -1  'True
         Caption         =   "&Prepare Bookfind files"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   -66375
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5100
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CheckBox chkLoyalty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Supports loyalty club"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -73695
         TabIndex        =   65
         Top             =   4335
         Width           =   2220
      End
      Begin VB.CheckBox chkAggregatePOs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Use open P.O.s for generated"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -69645
         TabIndex        =   64
         Top             =   4320
         Width           =   2715
      End
      Begin VB.Frame Frame3 
         Caption         =   "Transfer discounts"
         ForeColor       =   &H8000000D&
         Height          =   1140
         Left            =   -74550
         TabIndex        =   59
         Top             =   3030
         Width           =   3675
         Begin VB.TextBox txtTransferDiscountAdjustment 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2820
            TabIndex        =   61
            Top             =   675
            Width           =   540
         End
         Begin VB.TextBox txtTransferDiscount 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2820
            TabIndex        =   60
            Top             =   285
            Width           =   540
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Transfer discount adjustment"
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   150
            TabIndex        =   63
            Top             =   705
            Width           =   2580
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Transfer discount"
            ForeColor       =   &H8000000D&
            Height          =   315
            Left            =   975
            TabIndex        =   62
            Top             =   330
            Width           =   1755
         End
      End
      Begin VB.CommandButton cmdLoadDefaults 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Load defaults"
         Height          =   345
         Left            =   -69960
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1020
         Width           =   1650
      End
      Begin VB.ComboBox cboWorkstations 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -72795
         TabIndex        =   56
         Top             =   1020
         Width           =   2835
      End
      Begin VB.CheckBox chkEnforceCOLRefs 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Enforce refs on C.O.L."
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -69060
         TabIndex        =   53
         ToolTipText     =   "Customer order lines must all have references"
         Top             =   4020
         Width           =   2130
      End
      Begin VB.CheckBox chkReorderStyle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Reorder per C.O.L."
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -68820
         TabIndex        =   52
         ToolTipText     =   "The reordering slate allocates a purchase order line to each customer order line, with references copied across to purchase order."
         Top             =   3720
         Width           =   1890
      End
      Begin VB.TextBox txtCOLAllocStyle 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -67320
         TabIndex        =   51
         ToolTipText     =   "Retail environment holds stock on receipt and waits for customer to visit. Suppliers invoice immediately"
         Top             =   4680
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Frame Frame2 
         Caption         =   "Customer order types supported"
         ForeColor       =   &H8000000D&
         Height          =   1665
         Left            =   -74550
         TabIndex        =   45
         Top             =   1140
         Width           =   3660
         Begin VB.OptionButton optNeither 
            Caption         =   "Neither"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   285
            TabIndex        =   49
            Top             =   1260
            Width           =   2385
         End
         Begin VB.OptionButton optCOBoth 
            Caption         =   "Both above"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   270
            TabIndex        =   48
            Top             =   955
            Width           =   2385
         End
         Begin VB.OptionButton optCOStandingOrders 
            Caption         =   "Standing orders"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   270
            TabIndex        =   47
            Top             =   650
            Width           =   2385
         End
         Begin VB.OptionButton optCOWants 
            Caption         =   "Wants"
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   270
            TabIndex        =   46
            Top             =   345
            Width           =   2385
         End
      End
      Begin VB.CheckBox chkNonBookYN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Non-book capture"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -68880
         TabIndex        =   44
         Top             =   5520
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.CheckBox chkSupportsWORD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Supports WORD"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -70050
         TabIndex        =   43
         Top             =   5490
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CommandButton cmdEdiTsTAFF 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1470
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   5130
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddstaff 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5130
         Width           =   1095
      End
      Begin VB.CheckBox chkSections 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Enforce sections"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -68685
         TabIndex        =   39
         Top             =   3405
         Width           =   1755
      End
      Begin VB.CheckBox chkDiscountVATDefault 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Discount VAT default"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -73755
         TabIndex        =   38
         Top             =   4020
         Width           =   2280
      End
      Begin VB.CommandButton cmdLocateCS 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Locate cash sales customer"
         Height          =   375
         Left            =   -69840
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1050
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.CheckBox chkAntiquarian 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Configure for antiquarian usage"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -74445
         TabIndex        =   35
         Top             =   3720
         Width           =   2970
      End
      Begin VB.CheckBox chkCaptureDecimal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Capture using decimal in currencies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -70830
         TabIndex        =   34
         Top             =   5550
         Visible         =   0   'False
         Width           =   3465
      End
      Begin VB.TextBox txtLookupSeq 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -67860
         TabIndex        =   32
         Top             =   2790
         Width           =   915
      End
      Begin VB.ComboBox cboLocalCountry 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69765
         TabIndex        =   30
         Top             =   1905
         Width           =   2835
      End
      Begin VB.CommandButton cmdLocal 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set as local"
         Height          =   345
         Left            =   -66840
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1635
         Width           =   1485
      End
      Begin VB.CheckBox chkCC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Casual customers"
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -73485
         TabIndex        =   28
         Top             =   3405
         Width           =   2010
      End
      Begin VB.CommandButton cmdAddStore 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   405
         Left            =   -74745
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5130
         Width           =   1095
      End
      Begin VB.CommandButton cmdStoreEdit 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         Height          =   405
         Left            =   -73635
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5130
         Width           =   1095
      End
      Begin VB.CommandButton cmdSetDefaultStore 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set as default"
         Height          =   420
         Left            =   -68205
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1350
         Width           =   1500
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         Height          =   405
         Left            =   -72525
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5130
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1695
         Left            =   -74490
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   3720
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.CommandButton cmdAddCurr 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   405
         Left            =   -74775
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5010
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditCurr 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         Height          =   405
         Left            =   -73665
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5010
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefaultCurr 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set as default"
         Height          =   345
         Left            =   -66840
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1275
         Width           =   1485
      End
      Begin VB.CommandButton cmdRemCurr 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         Height          =   405
         Left            =   -72555
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5010
         Width           =   1095
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Remove"
         Height          =   390
         Left            =   -72555
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5055
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefault 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Set as default"
         Height          =   345
         Left            =   -68010
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1380
         Width           =   1470
      End
      Begin VB.CommandButton cmdEditComp 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Edit"
         Height          =   390
         Left            =   -73650
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5055
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddComp 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   390
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5055
         Width           =   1095
      End
      Begin VB.TextBox txtPrefix 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73020
         TabIndex        =   8
         Top             =   1065
         Width           =   915
      End
      Begin VB.TextBox txtVATRate 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   5
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73020
         TabIndex        =   7
         Top             =   1485
         Width           =   915
      End
      Begin VB.TextBox txtOfferSignature 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73035
         TabIndex        =   6
         Top             =   1905
         Width           =   2655
      End
      Begin VB.CheckBox chkSignTransactions 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Enforce signing of all documents"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -69810
         TabIndex        =   5
         Top             =   5520
         Visible         =   0   'False
         Width           =   3180
      End
      Begin VB.CheckBox chkAllowCopy 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Allow storage of copy information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   -70350
         TabIndex        =   4
         Top             =   5520
         Visible         =   0   'False
         Width           =   3195
      End
      Begin MSComctlLib.ListView lvwCompanies 
         Height          =   3615
         Left            =   -74775
         TabIndex        =   16
         Top             =   1320
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6376
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Default company"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwCurrencies 
         Height          =   3675
         Left            =   -74790
         TabIndex        =   21
         Top             =   1260
         Width           =   7830
         _ExtentX        =   13811
         _ExtentY        =   6482
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Symbol"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Format string"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Factor"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Default"
            Object.Width           =   1834
         EndProperty
      End
      Begin MSComctlLib.ListView lvwStores 
         Height          =   3735
         Left            =   -74775
         TabIndex        =   27
         Top             =   1335
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   6588
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Store role"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvwStaff 
         Height          =   3705
         Left            =   225
         TabIndex        =   42
         Top             =   1320
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   6535
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Short name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Phone"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cell"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Level"
            Object.Width           =   0
         EndProperty
      End
      Begin TrueOleDBGrid60.TDBDropDown DDPrinters 
         Height          =   1020
         Left            =   -74595
         OleObjectBlob   =   "frmConfiguration.frx":0CB5
         TabIndex        =   54
         Top             =   3570
         Width           =   4095
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   4185
         Left            =   -74745
         OleObjectBlob   =   "frmConfiguration.frx":32DE
         TabIndex        =   55
         Top             =   1530
         Width           =   6780
      End
      Begin VB.Label Label41 
         Caption         =   "Printers usually available (grey indicates not presently connected)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   -67860
         TabIndex        =   138
         Top             =   1110
         Width           =   2625
      End
      Begin VB.Label Label40 
         Caption         =   "Quotation text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -69915
         TabIndex        =   136
         Top             =   2385
         Width           =   1860
      End
      Begin VB.Label Label39 
         Caption         =   "Appro text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -69960
         TabIndex        =   134
         Top             =   1005
         Width           =   1860
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmConfiguration.frx":8368
         ForeColor       =   &H80000002&
         Height          =   915
         Left            =   -74790
         TabIndex        =   132
         Top             =   4995
         Width           =   6750
      End
      Begin VB.Label Label37 
         Caption         =   "Invoice/Credit note text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74535
         TabIndex        =   131
         Top             =   2385
         Width           =   1860
      End
      Begin VB.Label Label36 
         Caption         =   "Purchase order text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74520
         TabIndex        =   130
         Top             =   990
         Width           =   1860
      End
      Begin VB.Label Label35 
         Caption         =   "Statement message"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74550
         TabIndex        =   126
         Top             =   3420
         Width           =   1860
      End
      Begin VB.Label Label34 
         Caption         =   "Unused at present"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -69525
         TabIndex        =   125
         Top             =   3420
         Width           =   1860
      End
      Begin VB.Label Label16 
         Caption         =   "Purchase order text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74550
         TabIndex        =   124
         Top             =   945
         Width           =   1860
      End
      Begin VB.Label Label21 
         Caption         =   "Sales order text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74550
         TabIndex        =   123
         Top             =   2190
         Width           =   1860
      End
      Begin VB.Label Label31 
         Caption         =   "Quotation text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -69525
         TabIndex        =   122
         Top             =   945
         Width           =   1860
      End
      Begin VB.Label Label32 
         Caption         =   "Invoice text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -69525
         TabIndex        =   121
         Top             =   2190
         Width           =   1860
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmConfiguration.frx":843A
         ForeColor       =   &H80000002&
         Height          =   915
         Left            =   -74745
         TabIndex        =   118
         Top             =   4665
         Width           =   6750
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Default section type"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74640
         TabIndex        =   108
         Top             =   2130
         Width           =   2190
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Default product type"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74655
         TabIndex        =   106
         Top             =   1635
         Width           =   2190
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Default customer type"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74655
         TabIndex        =   104
         Top             =   1155
         Width           =   2190
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Launch IG"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -68580
         TabIndex        =   102
         Top             =   1035
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Promotion IG"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -68580
         TabIndex        =   101
         Top             =   1500
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Sale IG"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -68580
         TabIndex        =   100
         Top             =   1980
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Lunch IG"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -68580
         TabIndex        =   99
         Top             =   2445
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Minimum markup %"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -69810
         TabIndex        =   93
         Top             =   2385
         Width           =   1845
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "General ledger reference"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -69420
         TabIndex        =   81
         Top             =   2310
         Width           =   1965
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Contra account number for suppliers' invoices (Creditors)"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -71565
         TabIndex        =   79
         Top             =   1830
         Width           =   4110
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Contra account number for our invoices (Debtors)"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -71550
         TabIndex        =   77
         Top             =   1350
         Width           =   4095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "EDI ID number"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74850
         TabIndex        =   68
         Top             =   2340
         Width           =   1755
      End
      Begin VB.Label Label9 
         Caption         =   "Select workstation to edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -74685
         TabIndex        =   57
         Top             =   1065
         Width           =   2295
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer order allocation style ('R'etail or 'S'upplier)"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -71850
         TabIndex        =   50
         Top             =   4725
         Visible         =   0   'False
         Width           =   4350
      End
      Begin VB.Label lblCSCustomer 
         Caption         =   "Missing"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -66840
         TabIndex        =   37
         Top             =   1125
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Book lookup sequence (e.g. BF or BFWH or WH or WHBF). The books are looked for on the sources in the sequence indicated."
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   -73845
         TabIndex        =   33
         Top             =   2745
         Width           =   5880
      End
      Begin VB.Label Label2 
         Caption         =   "Local country"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -69750
         TabIndex        =   31
         Top             =   1695
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Transaction prefix"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74730
         TabIndex        =   11
         Top             =   1125
         Width           =   1635
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "VAT Rate"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74850
         TabIndex        =   10
         Top             =   1545
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Offer signature"
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   -74850
         TabIndex        =   9
         Top             =   1965
         Width           =   1755
      End
   End
   Begin VB.Label lblErrors 
      BackColor       =   &H00D3D3CB&
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   3930
      TabIndex        =   2
      Top             =   6360
      Width           =   1875
   End
End
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oConfig As a_Configuration
Dim flgLoading As Boolean
Dim tlOperators As New z_TextList
Dim tlCountries As New z_TextList
Dim tlWorkstations As New z_TextList
Dim tlIGs As New z_TextList
Dim lngOperatorID As Long
Dim XDOC As XArrayDB
Dim XPR As XArrayDB

Private Sub EnableOK(pOK As Boolean)
    On Error GoTo errHandler
    Me.cmdOK.Enabled = pOK
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.EnableOK(pOK)", pOK
End Sub
Private Sub oConfig_Valid(pErrors As String, Status As Boolean)
    On Error GoTo errHandler
    EnableOK Status
    lblErrors = pErrors
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.oConfig_Valid(pErrors,Status)", Array(pErrors, Status), EA_NORERAISE
    HandleError
End Sub

Public Sub component(poConfig As a_Configuration)
    On Error GoTo errHandler
    Set oConfig = poConfig
    oConfig.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.component(poConfig)", poConfig
End Sub
Private Sub RefreshData()
    On Error GoTo errHandler

    Screen.MousePointer = vbHourglass
    oPC.ReloadConfiguration
    LoadControls
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.RefreshData"
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim rs As New ADODB.Recordset
    flgLoading = True
    
    If oConfig.CSCustomerID > 0 Then Me.lblCSCustomer.Caption = "Found"
    Me.txtPOText = oConfig.OrderText
    Me.txtPrefix = oConfig.TransactionPrefix
    Me.txtVATRate = oConfig.VATRate
    Me.txtOfferSignature = oConfig.OfferSignature
    Me.txtLookupSeq = oConfig.LookupSeq
    txtTransferDiscount = oConfig.TFRDiscountF
    txtTransferDiscountAdjustment = oConfig.TFRDiscountAdjF
    Me.chkSignTransactions = IIf(oConfig.SignTransactions, 1, 0)
    Me.chkCaptureDecimal = IIf(oConfig.CaptureDecimal, 1, 0)
    Me.chkSections = IIf(oConfig.EnforceSections, 1, 0)
    Me.chkAllowCopy = IIf(oConfig.AllowCopyInfo, 1, 0)
    Me.chkAggregatePOs = IIf(oConfig.AggregatePOs, 1, 0)
    Me.chkAntiquarian = IIf(oConfig.AntiquarianYN, 1, 0)
    Me.chkReorderStyle = IIf(oConfig.ReorderPerCOL, 1, 0)
    Me.chkEnforceCOLRefs = IIf(oConfig.EnforceCOLRef, 1, 0)
    Me.chkNonBookYN = IIf(oConfig.NonBookYN, 1, 0)
    Me.chkSupportsWORD = IIf(oConfig.SupportsWORD, 1, 0)
    Me.chkCC = IIf(oConfig.CasualCustomersYN, 1, 0)
'    Me.txtUpdateStart = oConfig.UpdateWindowStartFormatted
'    Me.txtUpdateEnd = oConfig.UpdateWindowEndFormatted
  '  Me.txtCOLAllocStyle = oConfig.COLAllocationStyle
    Me.txtEDINumber = oConfig.GFXNumber
    Me.txtMU = oConfig.MinMU
    Me.txtStartLoy = oConfig.LoyaltyNumberingStartAt
    Me.txtEndLoy = oConfig.LoyaltyNumberingEndAt
    Me.txtEmailApproMsg = oConfig.EmailAPPMsg
    Me.txtEmailInvoiceMsg = oConfig.EmailInvMsg
    Me.txtEMailPOMsg = oConfig.EmailPOMsg
    Me.txtEmailQuotationMsg = oConfig.EmailQuoteMsg
    
    Select Case oConfig.AccountingApplicationName
    Case "NONE", ""
        optNone.Value = True
    Case "ACCPAC"
        optAccpac.Value = True
    Case "PASTEL"
        optPastel.Value = True
    End Select
    txtDebtorsContra = oConfig.DebtorsContraAccount
    txtCreditorsContra = oConfig.CreditorsContraAccount
    txtGLReference = oConfig.GLReference
    Me.txtCOText = oConfig.SalesOrderText
    Me.txtPOText = oConfig.OrderText
    Me.txtQUText = oConfig.QuotationText
    Me.txtInvText = oConfig.InvoiceText
    Me.txtStatementMessage = oConfig.StatementText
    Select Case oConfig.COTypesSupported
    Case 0
        Me.optNeither = True
    Case 1
        Me.optCOWants = True
    Case 2
        Me.optCOStandingOrders = True
    Case 3
        Me.optCOBoth = True
    End Select
    
    tlCountries.Load ltCountry
    LoadCombo Me.cboLocalCountry, tlCountries
    cboLocalCountry = tlCountries.Item(CStr(oConfig.LocalCountryID))
    
    LoadCombo Me.cboWorkstations, oPC.Configuration.Workstations
    cboWorkstations = oPC.WorkstationName
    LoadDocGrid oPC.Configuration.Workstations.Key(cboWorkstations)
    
    LoadCombo cboDefCustomerType, oPC.Configuration.CustomerTypesActive
    cboDefCustomerType = oPC.Configuration.CustomerTypesActive.Item(oPC.Configuration.DefaultCT)
    
    LoadCombo Me.cboDefPT, oPC.Configuration.ProductTypes
    cboDefPT = oPC.Configuration.ProductTypes.Item(oPC.Configuration.DefaultPT)
    
    LoadCombo cboDefCategory, oPC.Configuration.Sections
    Me.cboDefCategory = oPC.Configuration.Sections.Item(oPC.Configuration.DefaultSection)
    FillStoresList
    FillCompanyList
    FillCurrencyList
    FillStaffList
    FillPrintersList
    LoadDocDDPrinters
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.LoadControls"
End Sub

Private Sub FillPrintersList()
    On Error GoTo errHandler
Dim itmList As ListItem
Dim lngIndex As Long
    Me.lvwPrinters.ListItems.Clear
    For lngIndex = 1 To oConfig.Printers.Count
        Set itmList = lvwPrinters.ListItems.Add(Key:=oConfig.Printers.Key(oConfig.Printers.ItemByOrdinalIndex(lngIndex)) & "k")
        With itmList
            .text = oConfig.Printers.ItemByOrdinalIndex(lngIndex)
            If oConfig.Printers.ActiveByOrdinal(lngIndex) = False Then
                .ForeColor = COLOR_CANCELLED
            End If
        End With
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.FillPrintersList"
End Sub

Private Sub FillStoresList()
    On Error GoTo errHandler
Dim objItem As a_Store
Dim itmList As ListItem
Dim lngIndex As Long
    Me.lvwStores.ListItems.Clear
    For lngIndex = 1 To oConfig.Stores.Count
        With objItem
            Set objItem = oConfig.Stores.Item(lngIndex)
            Set itmList = lvwStores.ListItems.Add(Key:=objItem.Key)
            With itmList
                .text = objItem.Description
                If objItem.IsDeleted Then .text = .text & "(deleted)"
                If objItem.IsNew Then .text = .text & "(New)"
                .SubItems(1) = objItem.code
                If objItem.ID = oConfig.BillToStoreID Then .SubItems(2) = "bill to"
                If objItem.ID = oConfig.DefaultStoreID Then .SubItems(2) = "default"
            End With
        End With
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.FillStoresList"
End Sub
Private Sub FillStaffList()
    On Error GoTo errHandler
Dim objItem As a_Staff
Dim itmList As ListItem
Dim lngIndex As Long
    Me.lvwStaff.ListItems.Clear
    For lngIndex = 1 To oConfig.Staff.Count
        With objItem
            Set objItem = oConfig.Staff.Item(lngIndex)
            Set itmList = lvwStaff.ListItems.Add(Key:=objItem.Key)
            With itmList
                .text = objItem.StaffName
                If objItem.IsDeleted Then .text = .text & "(deleted)"
                If objItem.IsNew Then .text = .text & "(New)"
                If objItem.IsActive = False Then
                    itmList.ForeColor = vbRed
                End If
                .SubItems(1) = objItem.Shortname
                .SubItems(2) = objItem.StaffTel
                .SubItems(3) = objItem.StaffCell
               ' .SubItems(4) = objItem.Level
            End With
        End With
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.FillStaffList"
End Sub

Private Sub FillCompanyList()
    On Error GoTo errHandler
Dim objItem As a_Company
Dim itmList As ListItem
Dim lngIndex As Long

    Me.lvwCompanies.ListItems.Clear
    For lngIndex = 1 To oConfig.Companies.Count
        With objItem
            Set objItem = oConfig.Companies.Item(lngIndex)
            Set itmList = lvwCompanies.ListItems.Add(Key:=CStr(objItem.Key))
            With itmList
                .text = objItem.CompanyName
                If objItem.IsDeleted Then .text = .text & "(deleted)"
                If objItem.IsNew Then .text = .text & "(New)"
                .SubItems(1) = objItem.CompanyCode
                If objItem.ID = oConfig.DefaultCOMPID Then .SubItems(2) = "default"
            End With
        End With
    Next


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.FillCompanyList"
End Sub
Private Sub FillCurrencyList()
    On Error GoTo errHandler
Dim objItem As a_Currency
Dim itmList As ListItem
Dim lngIndex As Long

    lvwCurrencies.ListItems.Clear
    For lngIndex = 1 To oConfig.Currencies.Count
        With objItem
            Set objItem = oConfig.Currencies.Item(lngIndex)
            Set itmList = lvwCurrencies.ListItems.Add(Key:=objItem.Key)
            With itmList
                .text = objItem.Description
                If objItem.IsDeleted Then .text = .text & "(deleted)"
                If objItem.IsNew Then .text = .text & "(New)"
                .SubItems(1) = objItem.Symbol
                .SubItems(2) = objItem.FormatString
                .SubItems(3) = objItem.FactorF & "/" & objItem.FactorINVF
                If oConfig.DefaultCurrencyID = oConfig.LocalCurrencyID Then
                    If objItem.ID = oConfig.DefaultCurrencyID Then .SubItems(4) = "Default and Local"
                Else
                    If objItem.ID = oConfig.DefaultCurrencyID Then .SubItems(4) = "Default"
                    If objItem.ID = oConfig.LocalCurrencyID Then .SubItems(4) = "Local"
                End If
            '''''    If objItem.ID = oConfig.LocalCurrencyID Then .SubItems(2) = "default"
            End With
        End With
    Next


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.FillCurrencyList"
End Sub



Private Sub cboDefCustomerType_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oPC.Configuration.DefaultCT = oPC.Configuration.CustomerTypesActive.Key(Me.cboDefCustomerType)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cboDefCustomerType_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cboDefPT_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oPC.Configuration.DefaultPT = oPC.Configuration.ProductTypes.Key(Me.cboDefPT)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cboDefPT_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cboDefCategory_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oPC.Configuration.DefaultSection = oPC.Configuration.Sections.Key(Me.cboDefCategory)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cboDefCategory_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub cboLocalCountry_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oConfig.LocalCountryID = tlCountries.Key(cboLocalCountry)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cboLocalCountry_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cboWorkstations_Click()
    On Error GoTo errHandler
    LoadDocGrid oPC.Configuration.Workstations.Key(cboWorkstations)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cboWorkstations_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboWorkstations_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    LoadDocGrid oPC.Configuration.Workstations.Key(cboWorkstations)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cboWorkstations_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub chkAggregatePOs_Click()
    On Error GoTo errHandler
    oConfig.AggregatePOs = IIf(chkAggregatePOs = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkAggregatePOs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkAllowCopy_Click()
    On Error GoTo errHandler
    oConfig.AllowCopyInfo = IIf(chkAllowCopy = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkAllowCopy_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkCaptureDecimal_Click()
    On Error GoTo errHandler
    oConfig.CaptureDecimal = IIf(chkCaptureDecimal = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkCaptureDecimal_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub chkAntiquarian_Click()
    On Error GoTo errHandler
    oConfig.AntiquarianYN = IIf(chkAntiquarian = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkAntiquarian_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkCC_Click()
    On Error GoTo errHandler
    oConfig.CasualCustomersYN = IIf(chkCC = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkCC_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkDiscountVATDefault_Click()
    On Error GoTo errHandler
    oConfig.DiscountVATDefault = IIf(chkDiscountVATDefault = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkDiscountVATDefault_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkEnforceCOLRefs_Click()
    On Error GoTo errHandler
    oConfig.EnforceCOLRef = IIf(chkEnforceCOLRefs = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkEnforceCOLRefs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkLoyalty_Click()
    On Error GoTo errHandler
    oConfig.SupportsLoyaltyClub = IIf(chkLoyalty = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkLoyalty_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkNonBookYN_Click()
    On Error GoTo errHandler
    oConfig.NonBookYN = IIf(chkNonBookYN = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkNonBookYN_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkReorderStyle_Click()
    On Error GoTo errHandler
    oConfig.ReorderPerCOL = IIf(chkReorderStyle = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkReorderStyle_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub chkBookfind_Click()
'    oConfig.UsesBookfind = IIf(chkBookfind = 1, True, False)
'End Sub
'
'Private Sub chkGardners_Click()
'    oConfig.UsesGardners = IIf(chkGardners = 1, True, False)
'End Sub

'Private Sub chkPI_Click()
'    oConfig.ShowProdsWithInstancesOnly = IIf(chkPI = 1, True, False)
'End Sub

Private Sub chkSections_Click()
    On Error GoTo errHandler
    oConfig.EnforceSections = IIf(chkSections = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkSections_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkSupportsWORD_Click()
    On Error GoTo errHandler
    oConfig.SupportsWORD = IIf(chkSupportsWORD = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkSupportsWORD_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkSignTransactions_Click()
    On Error GoTo errHandler
    oConfig.SignTransactions = IIf(chkSignTransactions = 1, True, False)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.chkSignTransactions_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdAddComp_Click()
    On Error GoTo errHandler
Dim oComp As a_Company
Dim frm As frmCompany

    Set oComp = oConfig.Companies.Add
 '   oCOmp.BeginEdit
    Set frm = New frmCompany
    frm.component oComp
    frm.Show vbModal
    FillCompanyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdAddComp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddCurr_Click()
    On Error GoTo errHandler
Dim oCurr As a_Currency
Dim frm As frmCurrency

    Set oCurr = oConfig.Currencies.Add
    Set frm = New frmCurrency
    frm.component oCurr
    frm.Show vbModal
    FillCurrencyList

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdAddCurr_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdAddStore_Click()
    On Error GoTo errHandler
Dim oStore As a_Store
Dim frm As frmStore

    Set oStore = oConfig.Stores.Add
    Set frm = New frmStore
    frm.component oStore
    frm.Show vbModal
    FillStoresList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdAddStore_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdBillToStore_Click()
    On Error GoTo errHandler
    oConfig.BillToStoreID = oConfig.Stores.FindStoreByID(val(lvwStores.SelectedItem.Key)).ID
    FillStoresList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdBillToStore_Click"
    HandleError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If MsgBox("Any alterations you have made to configuration will be left unchanged. ", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    oConfig.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdCleanup_Click()
'    On Error GoTo errHandler
'    LoadListbox Me.lbPrinters, oPC.Configuration.Printers
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmConfiguration.cmdCleanup_Click", , EA_NORERAISE
'    HandleError
'End Sub

'Private Sub cmdCopyBFFiles_Click()
'Dim oFSO As New FileSystemObject
'Dim fold
'Dim fc
'Dim f
''    If Not oPC.BFLoaded Then
'        fold = oFSO.GetFolder(oPC.BookFindRoot
'    Else
'        MsgBox "This operation can only work if you are not presently connected to Bookfind from this application."
'    End If
'End Sub

Private Sub cmdDefault_Click()
    On Error GoTo errHandler
    oConfig.DefaultCOMPID = oConfig.Companies(val(lvwCompanies.SelectedItem.Key)).ID
    FillCompanyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdDefault_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDefaultCurr_Click()
    On Error GoTo errHandler
    oConfig.DefaultCurrencyID = oConfig.Currencies(lvwCurrencies.SelectedItem.Key).ID
    FillCurrencyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdDefaultCurr_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDeletePrinter_Click()
    On Error GoTo errHandler
    
    oPC.Configuration.DeletePrinter CLng(val(lvwPrinters.SelectedItem.Key))
    FillPrintersList
    LoadDocDDPrinters
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdDeletePrinter_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEditComp_Click()
    On Error GoTo errHandler
Dim frm As New frmCompany
Dim oComp As a_Company
Dim lngResult As Long
    Set oComp = New a_Company
   ' Set oComp = oConfig.Companies(CStr(lvwCompanies.SelectedItem.key))
    Set oComp = oConfig.Companies.Item(val(lvwCompanies.SelectedItem.Key))
    frm.component oComp
    frm.Show vbModal
    FillCompanyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdEditComp_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEditCurr_Click()
    On Error GoTo errHandler
    If Not SecurityControlforSupervisor Then
        Exit Sub
    End If
    EditCurr
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdEditCurr_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdEditStaff_Click()
    On Error GoTo errHandler

    EditStaff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdEditStaff_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub EditStaff()
    On Error GoTo errHandler
Dim frm As New frmStaff
Dim oStaff As a_Staff
Dim lngResult As Long

    If lvwStaff.SelectedItem.text = "OLD_STAFF" Then
        MsgBox "You cannot edit this item.", vbOKOnly, "Can't do this"
        Exit Sub
    End If

    If lvwStaff.SelectedItem Is Nothing Then Exit Sub
        Set oStaff = New a_Staff
        Set oStaff = oConfig.Staff(lvwStaff.SelectedItem.Key)
        frm.component oStaff
        frm.Show vbModal
        FillStaffList
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.EditStaff"
End Sub
Private Sub cmdAddstaff_Click()
    On Error GoTo errHandler
Dim oStaff As a_Staff
Dim frm As frmStaff

    If SecurityControlforSupervisor Then
        Set oStaff = oConfig.Staff.Add
        Set frm = New frmStaff
        frm.component oStaff
        frm.Show vbModal
        FillStaffList
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdAddstaff_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLocal_Click()
    On Error GoTo errHandler
    oConfig.LocalCurrencyID = oConfig.Currencies(lvwCurrencies.SelectedItem.Key).ID
    FillCurrencyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdLocal_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdLocateCS_Click()
    On Error GoTo errHandler
Dim rs As ADODB.Recordset

Dim OpenResult As Integer
    '--------------
    OpenResult = oPC.OpenDBSHort
    '--------------
    Set rs = New ADODB.Recordset
    rs.open "SELECT TP_NAME,TP_ID,TP_ACNO FROM tTP WHERE TP_ACNO = '" & "CASH01" & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
    If Not rs.eof Then
        oConfig.CSCustomerID = rs!TP_ID
        lblCSCustomer.Caption = "LOCATED"
    Else
        MsgBox "You must create a customer with Acc. no. 'CASH01' in order to process cash sales." & vbCrLf & "When you have done so, return to this point and click the button again.", , "Warning"
        lblCSCustomer.Caption = "MISSING"
    End If
'    --------------
    If OpenResult = 0 Then oPC.DisconnectDBShort  'if the recent open command actually opened a connection then close it
'    --------------
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdLocateCS_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim strStatus As String
    Screen.MousePointer = vbHourglass
    oConfig.ApplyEdit strStatus
    Screen.MousePointer = vbDefault
    If strStatus <> "" Then
        strStatus = "The save operation has not been successful for the following reason:" & vbCrLf & vbCrLf & strStatus & vbCrLf & vbCrLf & "Either select Cancel or correct the data and select OK again."
        MsgBox strStatus
    Else
        Unload Me
        Screen.MousePointer = vbHourglass
        oPC.ReloadConfiguration
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo errHandler
    RefreshData
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdRefresh_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemCurr_Click()
    On Error GoTo errHandler
Dim oCurr As a_Currency
Dim lngResult As Long
Dim oSQL As New z_SQL
Dim bCanDelete As Boolean

    bCanDelete = (oSQL.QtyDocumentsUsingCurrency(oConfig.Currencies.Item(lvwCurrencies.SelectedItem.Key).ID) = 0)
    If bCanDelete Then
        oConfig.Currencies.Item(lvwCurrencies.SelectedItem.Key).Delete
        FillCurrencyList
    Else
        MsgBox "This currency is associated with documents in your database. You cannot delete it." & vbCrLf _
        & "You should first merge this currency with another, then you will be able to delete it.", vbInformation + vbOKOnly, "Can't do this"
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdRemCurr_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo errHandler
    oConfig.Companies.Remove (lvwCompanies.SelectedItem.Index)
    FillCompanyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdRemove_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemoveStaff_Click()
    On Error GoTo errHandler
Dim oUtil As z_UTIL
    If lvwStaff.SelectedItem.text = "OLD_STAFF" Then
        MsgBox "You cannot remove this item.", vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If MsgBox("You want to remove " & lvwStaff.SelectedItem.text & " from the list of staffmembers?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Set oUtil = New z_UTIL
    oUtil.RemoveStaffmember oConfig.Staff.Item(lvwStaff.SelectedItem.Key).ID
    lvwStaff.SelectedItem.text = lvwStaff.SelectedItem.text & "(REMOVED)"
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdRemoveStaff_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRolldates_Click()
    On Error GoTo errHandler
    Me.dp90 = Me.dp60
    Me.dp60 = Me.dp30
    Me.dp30 = Me.dpCur
    Me.dpCur = Date
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdRolldates_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdRemoveDC_Click()
'    oConfig.DocumentControl.Remove (lvwDC.SelectedItem.Index)
'    FillDCList
'End Sub

Private Sub cmdSetDefaultStore_Click()
    On Error GoTo errHandler
    oConfig.DefaultStoreID = oConfig.Stores.Item(val(lvwStores.SelectedItem.Key)).ID
    If oConfig.BillToStoreID = 0 Then oConfig.BillToStoreID = oConfig.DefaultStoreID
    FillStoresList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdSetDefaultStore_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdStoreEdit_Click()
    On Error GoTo errHandler
    EditStore
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdStoreEdit_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdTestErrorHandling_DeveloperOnly_Click()
    On Error GoTo errHandler
    Err.Raise 13243, "test", "Test Error Message"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdTestErrorHandling_DeveloperOnly_Click", , EA_NORERAISE
    HandleError
End Sub



'Private Sub Combo1_Click()
'    oConfig.OpSetsAuto = tlOperators.Key(Combo1.Text)
'End Sub
'
'
'
'Private Sub EditCurr_Click()
'
'End Sub
'
'
'
'Private Sub Command2_Click()
'
'End Sub
'
Private Sub Form_Load()
    On Error GoTo errHandler
Me.Width = 10700
Me.Height = 7500
Me.TOP = 400
Me.Left = 500
    LoadControls
    Me.SSTab1.Tab = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.Form_Load", , EA_NORERAISE
    HandleError
End Sub




Private Sub lvwCompanies_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwCompanies_AfterLabelEdit(Cancel,NewString)", Array(Cancel, _
         NewString), EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCompanies_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwCompanies_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCompanies_DblClick()
    On Error GoTo errHandler
Dim frm As New frmCompany
Dim oComp As a_Company
Dim lngResult As Long
    If lvwCompanies.SelectedItem.Index < 1 Then Exit Sub

    Set oComp = New a_Company
    Set oComp = oConfig.Companies.Item(val(lvwCompanies.SelectedItem.Key))
  '  oCOmp.BeginEdit
    frm.component oComp
    frm.Show vbModal
    FillCompanyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwCompanies_DblClick", , EA_NORERAISE
    HandleError
End Sub



Private Sub lvwCurrencies_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwCurrencies_AfterLabelEdit(Cancel,NewString)", Array(Cancel, _
         NewString), EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCurrencies_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwCurrencies_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwCurrencies_DblClick()
    On Error GoTo errHandler
    If Not SecurityControlforSupervisor Then
        Exit Sub
    End If
    EditCurr
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwCurrencies_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub EditCurr()
    On Error GoTo errHandler
Dim frm As New frmCurrency
Dim oCurr As a_Currency
Dim lngResult As Long

    Set oCurr = New a_Currency
    Set oCurr = oConfig.Currencies.Item(val(lvwCurrencies.SelectedItem.Key))
    frm.component oCurr
    frm.Show vbModal
    FillCurrencyList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.EditCurr"
End Sub
Private Sub EditStore()
    On Error GoTo errHandler
Dim frm As New frmStore
Dim oStore As a_Store
Dim lngResult As Long

    Set oStore = New a_Store
    Set oStore = oConfig.Stores.Item(lvwStores.SelectedItem.Key)
    frm.component oStore
    frm.Show vbModal
    FillStoresList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.EditStore"
End Sub

Private Sub lvwPrinters_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub lvwStaff_DblClick()
    On Error GoTo errHandler
    EditStaff
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwStaff_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub lvwStores_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwStores_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwStores_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwStores_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub lvwStaff_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwStaff_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwDC_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwDC_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub lvwDC_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwDC_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwStaff_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwStaff_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwStores_DblClick()
    On Error GoTo errHandler
    EditStore
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.lvwStores_DblClick", , EA_NORERAISE
    HandleError
End Sub



Private Sub optCOBoth_Click()
    On Error GoTo errHandler
   If flgLoading Then Exit Sub
    If optCOBoth Then
        oConfig.COTypesSupported = 3
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.optCOBoth_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optCOStandingOrders_Click()
    On Error GoTo errHandler
   If flgLoading Then Exit Sub
    If optCOStandingOrders Then
        oConfig.COTypesSupported = 2
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.optCOStandingOrders_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optCOWants_Click()
    On Error GoTo errHandler
   If flgLoading Then Exit Sub
    If optCOWants Then
        oConfig.COTypesSupported = 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.optCOWants_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optNeither_Click()
    On Error GoTo errHandler
   If flgLoading Then Exit Sub
    If optNeither Then
        oConfig.COTypesSupported = 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.optNeither_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub optPastel_Click()
    On Error GoTo errHandler
    If optPastel = True Then
        oPC.Configuration.AccountingApplicationName = "PASTEL"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.optPastel_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub optAccpac_Click()
    On Error GoTo errHandler
    If optAccpac = True Then
        oPC.Configuration.AccountingApplicationName = "ACCPAC"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.optAccpac_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub optNone_Click()
    On Error GoTo errHandler
    If optNone = True Then
        oPC.Configuration.AccountingApplicationName = "NONE"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.optNone_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub txtCOLAllocStyle_Change()
'    On Error GoTo errHandler
'Dim intPos As Integer
'
'   If flgLoading Then Exit Sub
'    On Error Resume Next
'    oConfig.COLAllocationStyle = txtCOLAllocStyle
'    If Err Then
'      Beep
'      intPos = txtCOLAllocStyle.SelStart
'      txtCOLAllocStyle = oConfig.COLAllocationStyle
'      txtCOLAllocStyle.SelStart = intPos - 1
'    End If
'
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmConfiguration.txtCOLAllocStyle_Change", , EA_NORERAISE
'    HandleError
'End Sub



Private Sub txtDebtorsContra_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.DebtorsContraAccount = Trim(txtDebtorsContra)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtDebtorsContra_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCreditorsContra_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.CreditorsContraAccount = Trim(txtCreditorsContra)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtCreditorsContra_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub txtEndLoy_Change()
    On Error GoTo errHandler
    oConfig.SetLoyaltyNumberingEndAt txtEndLoy
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEndLoy_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtStartLoy_Change()
    On Error GoTo errHandler
    oConfig.SetLoyaltyNumberingStartAt txtStartLoy
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtStartLoy_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtGLReference_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oPC.Configuration.GLReference = Trim(txtGLReference)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtGLReference_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtLookupSeq_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.SetLookupSequence txtLookupSeq
    If Err Then
      Beep
      intPos = txtLookupSeq.SelStart
      txtLookupSeq = oConfig.LookupSeq
      txtLookupSeq.SelStart = intPos - 1
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtLookupSeq_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtLookupSeq_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtLookupSeq")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtLookupSeq_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtLookupSeq_LostFocus()
    On Error GoTo errHandler
   txtLookupSeq.text = oConfig.LookupSeq
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtLookupSeq_LostFocus", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtMU_LostFocus()
    On Error GoTo errHandler
  ' txtMU.Text = oConfig.MinMU
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtMU_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtMU_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If IsNumeric(txtMU) Then
        oConfig.MinMU = FNN(txtMU)
    Else
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtMU_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtOfferSignature_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.OfferSignature = txtOfferSignature
    If Err Then
      Beep
      intPos = txtOfferSignature.SelStart
      txtOfferSignature = oConfig.OfferSignature
      txtOfferSignature.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtOfferSignature_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtOfferSignature_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtOfferSignature")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtOfferSignature_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtOfferSignature_LostFocus()
    On Error GoTo errHandler
   txtOfferSignature.text = oConfig.OfferSignature
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtOfferSignature_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEDINumber_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.SetGFXNumber txtEDINumber
    If Err Then
      Beep
      intPos = txtEDINumber.SelStart
      txtEDINumber = oConfig.GFXNumber
      txtEDINumber.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEDINumber_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEDINumber_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtEDINumber")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEDINumber_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEDINumber_LostFocus()
    On Error GoTo errHandler
   txtEDINumber.text = oConfig.GFXNumber
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEDINumber_LostFocus", , EA_NORERAISE
    HandleError
End Sub




Private Sub txtPOText_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.OrderText = txtPOText
    If Err Then
      Beep
      intPos = txtPOText.SelStart
      txtPOText = oConfig.OrderText
      txtPOText.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtPOText_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPOText_LostFocus()
    On Error GoTo errHandler
   txtPOText.text = oConfig.OrderText
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtPOText_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmailPOMsg_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.EmailPOMsg = txtEMailPOMsg
    If Err Then
      Beep
      intPos = txtEMailPOMsg.SelStart
      txtEMailPOMsg = oConfig.EmailPOMsg
      txtEMailPOMsg.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEmailPOMsg_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmailPOMsg_LostFocus()
    On Error GoTo errHandler
   txtEMailPOMsg.text = oConfig.EmailPOMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEmailPOMsg_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmailInvoiceMsg_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.EmailInvMsg = txtEmailInvoiceMsg
    If Err Then
      Beep
      intPos = txtEmailInvoiceMsg.SelStart
      txtEmailInvoiceMsg = oConfig.EmailInvMsg
      txtEmailInvoiceMsg.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEmailInvoiceMsg_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmailInvoiceMsg_LostFocus()
    On Error GoTo errHandler
   txtEmailInvoiceMsg.text = oConfig.EmailInvMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEmailInvoiceMsg_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmailQuotationMsg_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.EmailQuoteMsg = txtEmailQuotationMsg
    If Err Then
      Beep
      intPos = txtEmailQuotationMsg.SelStart
      txtEmailQuotationMsg = oConfig.EmailQuoteMsg
      txtEmailQuotationMsg.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEmailQuotationMsg_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmailQuotationMsg_LostFocus()
    On Error GoTo errHandler
   txtEmailQuotationMsg.text = oConfig.EmailQuoteMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEmailQuotationMsg_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmailApproMsg_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.EmailAPPMsg = txtEmailApproMsg
    If Err Then
      Beep
      intPos = txtEmailApproMsg.SelStart
      txtEmailApproMsg = oConfig.EmailAPPMsg
      txtEmailApproMsg.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEmailApproMsg_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEmailApproMsg_LostFocus()
    On Error GoTo errHandler
   txtEmailApproMsg.text = oConfig.EmailAPPMsg
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtEmailApproMsg_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtStatementMessage_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.StatementText = txtStatementMessage
    If Err Then
      Beep
      intPos = txtStatementMessage.SelStart
      txtStatementMessage = oConfig.StatementText
      txtStatementMessage.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtStatementMessage_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtStatementMessage_LostFocus()
    On Error GoTo errHandler
   txtStatementMessage.text = oConfig.StatementText
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtStatementMessage_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCOText_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.SalesOrderText = txtCOText
    If Err Then
      Beep
      intPos = txtCOText.SelStart
      txtCOText = oConfig.OrderText
      txtCOText.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtCOText_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCOText_LostFocus()
    On Error GoTo errHandler
   txtCOText.text = oConfig.SalesOrderText
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtCOText_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQUText_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.QuotationText = txtQUText
    If Err Then
      Beep
      intPos = txtQUText.SelStart
      txtQUText = oConfig.OrderText
      txtQUText.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtQUText_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtQUText_LostFocus()
    On Error GoTo errHandler
   txtQUText.text = oConfig.QuotationText
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtQUText_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtINVText_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.InvoiceText = txtInvText
    If Err Then
      Beep
      intPos = txtInvText.SelStart
      txtInvText = oConfig.OrderText
      txtInvText.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtINVText_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtINVText_LostFocus()
    On Error GoTo errHandler
   txtInvText.text = oConfig.InvoiceText
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtINVText_LostFocus", , EA_NORERAISE
    HandleError
End Sub




Private Sub txtPrefix_Change()
    On Error GoTo errHandler
Dim intPos As Integer

   If flgLoading Then Exit Sub
    On Error Resume Next
    oConfig.TransactionPrefix = txtPrefix
    If Err Then
      Beep
      intPos = txtPrefix.SelStart
      txtPrefix = oConfig.TransactionPrefix
      txtPrefix.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtPrefix_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrefix_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtPrefix")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtPrefix_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrefix_LostFocus()
    On Error GoTo errHandler
   txtPrefix.text = oConfig.TransactionPrefix
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtPrefix_LostFocus", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtTransferDiscount_LostFocus()
    On Error GoTo errHandler
    txtTransferDiscount = oConfig.TFRDiscountF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtTransferDiscount_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTransferDiscount_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim iRes As Integer
    If ConvertToInt(txtTransferDiscount, iRes) Then
        oConfig.TFRDiscount = iRes
    Else
        oConfig.TFRDiscount = 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtTransferDiscount_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtTransferDiscountAdjustment_LostFocus()
    On Error GoTo errHandler
    txtTransferDiscountAdjustment = oConfig.TFRDiscountAdjF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtTransferDiscountAdjustment_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTransferDiscountAdjustment_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim iRes As Integer
    If ConvertToInt(txtTransferDiscountAdjustment, iRes) Then
        oConfig.TFRDiscountAdj = iRes
    Else
        oConfig.TFRDiscountAdj = 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtTransferDiscountAdjustment_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtPrintingSettings_Validate(Cancel As Boolean)
'    oConfig.SetPrintingSettings txtPrintingSettings
'End Sub

Private Sub txtUpdateStart_GotFocus()
    On Error GoTo errHandler
    AutoSelect Controls("txtUpdateStart")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtUpdateStart_GotFocus", , EA_NORERAISE
    HandleError
End Sub
'Private Sub txtUpdateStart_LostFocus()
'    txtUpdateStart = oConfig.UpdateWindowStartFormatted
'End Sub
'
'Private Sub txtUpdateStart_Validate(Cancel As Boolean)
'    Cancel = Not oConfig.SetUpdateWindowStart(txtUpdateStart)
'End Sub
'Private Sub txtUpdateEnd_GotFocus()
'    AutoSelect Controls("txtUpdateEnd")
'End Sub
'
'Private Sub txtUpdateEnd_Validate(Cancel As Boolean)
'    Cancel = Not oConfig.SetUpdateWindowEnd(txtUpdateEnd)
'End Sub
'Private Sub txtUpdateEnd_LostFocus()
'    txtUpdateEnd = oConfig.UpdateWindowEndFormatted
'End Sub

Private Sub txtVATRate_Change()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    
    oConfig.SetVATRate txtVATRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtVATRate_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtVATRate_LostFocus()
    On Error GoTo errHandler
   txtVATRate.text = oConfig.VATRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.txtVATRate_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdLoadDefaults_Click()
    On Error GoTo errHandler
    If MsgBox("This will cancel any changes to this form that you have not saved. Continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    DefaultLoad
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.cmdLoadDefaults_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub DefaultLoad()
    On Error GoTo errHandler
    oConfig.DocumentControls.VerifyDocumentTypes cboWorkstations
    oConfig.CancelEdit
    oConfig.Reload
    oConfig.BeginEdit
    LoadDocGrid oConfig.Workstations.Key(cboWorkstations)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.DefaultLoad"
End Sub
Private Sub LoadDocGrid(pWorkstationID As Long)
    On Error GoTo errHandler
Dim lngIndex As Long
Dim ArrayIdx As Long
Dim objItem As a_DocumentControl

    Set XDOC = New XArrayDB
    XDOC.Clear
    lngIndex = 1
    ArrayIdx = 1
    XDOC.ReDim 1, 0, 1, 8
    
    For lngIndex = 1 To oConfig.DocumentControls.Count
        If oConfig.DocumentControls(lngIndex).WorkstationID = pWorkstationID Then
            XDOC.ReDim 1, ArrayIdx, 1, 8
            XDOC.Value(ArrayIdx, 1) = oConfig.DocumentControls(lngIndex).DOCTypeName
          '  XDOC.Value(ArrayIdx, 2) = oConfig.DocumentControls(lngIndex).style
          '  XDOC.Value(ArrayIdx, 3) = oConfig.DocumentControls(lngIndex).PreviewPrintF
            XDOC.Value(ArrayIdx, 4) = oConfig.DocumentControls(lngIndex).PrinterName
            XDOC.Value(ArrayIdx, 5) = oConfig.DocumentControls(lngIndex).QtyCopies
            XDOC.Value(ArrayIdx, 6) = oConfig.DocumentControls(lngIndex).Key
            ArrayIdx = ArrayIdx + 1
        End If
    Next
    XDOC.QuickSort 1, ArrayIdx - 1, 1, XORDER_ASCEND, XTYPE_STRING
    Grid1.Array = XDOC
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.LoadDocGrid(pWorkstationID)", pWorkstationID
End Sub

Private Sub LoadDocDDPrinters()
    On Error GoTo errHandler
Dim lngIndex As Long
Dim ArrayIdx As Long
Dim objItem As a_DocumentControl
Dim vntItem As Variant

    Set XPR = New XArrayDB
    XPR.Clear
    ArrayIdx = 1
    XPR.ReDim 1, oConfig.Printers.Count, 1, 4
    

    For Each vntItem In oConfig.Printers
            XPR.Value(ArrayIdx, 1) = vntItem(0)
            XPR.Value(ArrayIdx, 2) = vntItem(1)
            ArrayIdx = ArrayIdx + 1
    Next
    XPR.QuickSort 1, ArrayIdx - 1, 1, XORDER_ASCEND, XTYPE_STRING
    DDPrinters.Array = XPR
    DDPrinters.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.LoadDocDDPrinters"
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim i As Integer
Dim oDC As a_DocumentControl
Dim lngResult As Long

    i = ColIndex + 1
    Set oDC = oConfig.DocumentControls(XDOC(Grid1.Bookmark, 6))
    oDC.BeginEdit
    Select Case i
    Case 2
        oDC.Style = Trim(Grid1.text)
    Case 3
        oDC.SetPPrintPreview FNS(Grid1.text)
    Case 4
        oDC.SetPrinter oConfig.Printers.Key(Trim(Grid1.text)), Trim(Grid1.text)
    Case 5
        If ConvertToLng(Grid1.text, lngResult) Then
            oDC.QtyCopies = lngResult
        End If
    End Select
    oDC.ApplyEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmConfiguration.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub



