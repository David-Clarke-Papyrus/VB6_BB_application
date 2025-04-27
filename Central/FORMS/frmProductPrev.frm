VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmProductPrev 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Stock"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   13980
   ControlBox      =   0   'False
   Icon            =   "frmProductPrev.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   13980
   Begin VB.CheckBox chkExSales 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Exclude from sales reporting"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   3585
      TabIndex        =   103
      Top             =   5520
      Width           =   2955
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   525
      Left            =   3870
      MultiLine       =   -1  'True
      TabIndex        =   95
      Top             =   5520
      Width           =   5175
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10380
      Picture         =   "frmProductPrev.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   5475
      Width           =   1110
   End
   Begin VB.CommandButton cmdEdit 
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
      Height          =   645
      Left            =   9270
      Picture         =   "frmProductPrev.frx":03B5
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   5475
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Find By ISBN"
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   240
      TabIndex        =   12
      Top             =   5415
      Width           =   3255
      Begin VB.CommandButton cmdsearchisbn 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   195
         Width           =   945
      End
      Begin VB.TextBox txtisbnsearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   150
         TabIndex        =   13
         Top             =   270
         Width           =   1995
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4005
      Left            =   180
      TabIndex        =   6
      Top             =   1335
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   535
      BackColor       =   14737632
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1. Stock"
      TabPicture(0)   =   "frmProductPrev.frx":06BF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label16"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label21"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label31"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label32"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblStatus"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label40"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "StGrid"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtOnHand"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtReserved"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtRRP"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSP"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCost"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtTotalSold"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtUKPrice"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtUSPrice"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtReturnable"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProductPrev.frx":06DB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "Label26"
      Tab(1).Control(3)=   "lblObsolete"
      Tab(1).Control(4)=   "lblDeal"
      Tab(1).Control(5)=   "lblSupplier"
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(7)=   "Label6"
      Tab(1).Control(8)=   "Label8"
      Tab(1).Control(9)=   "Label7"
      Tab(1).Control(10)=   "Label23"
      Tab(1).Control(11)=   "Label27"
      Tab(1).Control(12)=   "Label33"
      Tab(1).Control(13)=   "Label34"
      Tab(1).Control(14)=   "Label35"
      Tab(1).Control(15)=   "txtBinding"
      Tab(1).Control(16)=   "txtVAT"
      Tab(1).Control(17)=   "txtSection"
      Tab(1).Control(18)=   "txtPubPlace"
      Tab(1).Control(19)=   "txtPubDate"
      Tab(1).Control(20)=   "txtEdition"
      Tab(1).Control(21)=   "txtPublisher"
      Tab(1).Control(22)=   "txtSS"
      Tab(1).Control(23)=   "txtDefaultDeliveryDays"
      Tab(1).Control(24)=   "txtProductType"
      Tab(1).Control(25)=   "Frame3"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "&3. Notes"
      TabPicture(2)   =   "frmProductPrev.frx":06F7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label28"
      Tab(2).Control(1)=   "Label30"
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(4)=   "txtDescription"
      Tab(2).Control(5)=   "txtComment"
      Tab(2).Control(6)=   "txtFlagText"
      Tab(2).Control(7)=   "txtCategoryHeading"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "&4. Statistics"
      TabPicture(3)   =   "frmProductPrev.frx":0713
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label12"
      Tab(3).Control(1)=   "Label22"
      Tab(3).Control(2)=   "Label24"
      Tab(3).Control(3)=   "Label15"
      Tab(3).Control(4)=   "Label1"
      Tab(3).Control(5)=   "Label36"
      Tab(3).Control(6)=   "Label37"
      Tab(3).Control(7)=   "Label38"
      Tab(3).Control(8)=   "Label39"
      Tab(3).Control(9)=   "Frame2"
      Tab(3).Control(10)=   "txtDateAdded"
      Tab(3).Control(11)=   "txtDateLastModified"
      Tab(3).Control(12)=   "txtLastOrdered"
      Tab(3).Control(13)=   "txtLastOrderedQty"
      Tab(3).Control(14)=   "txtLastOrderedPrice"
      Tab(3).Control(15)=   "txtLastReceivedPrice"
      Tab(3).Control(16)=   "txtLastReceivedQty"
      Tab(3).Control(17)=   "txtLastReceived"
      Tab(3).Control(18)=   "txtLastCounted"
      Tab(3).Control(19)=   "txtLastCountedQty"
      Tab(3).Control(20)=   "txtLastCountedPrice"
      Tab(3).Control(21)=   "txtLastSoldPrice"
      Tab(3).Control(22)=   "txtLastSoldQty"
      Tab(3).Control(23)=   "txtLastSoldDate"
      Tab(3).ControlCount=   24
      TabCaption(4)   =   "&5. Copies"
      TabPicture(4)   =   "frmProductPrev.frx":072F
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Grid1"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&6. Sales"
      TabPicture(5)   =   "frmProductPrev.frx":074B
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label41"
      Tab(5).Control(1)=   "Chart2"
      Tab(5).Control(2)=   "chart1"
      Tab(5).Control(3)=   "cmdExpand"
      Tab(5).ControlCount=   4
      Begin VB.Frame Frame3 
         Caption         =   "BIC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1695
         Left            =   -74640
         TabIndex        =   100
         Top             =   1650
         Width           =   3690
         Begin VB.TextBox txtBIC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   330
            Left            =   165
            Locked          =   -1  'True
            TabIndex        =   102
            Top             =   360
            Width           =   2370
         End
         Begin VB.TextBox txtBICDescription 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
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
            Height          =   810
            Left            =   165
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   101
            Top             =   720
            Width           =   3360
         End
      End
      Begin VB.TextBox txtCategoryHeading 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   -74700
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   3480
         Width           =   10770
      End
      Begin VB.TextBox txtFlagText 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Left            =   -74715
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   2505
         Width           =   7215
      End
      Begin VB.TextBox txtReturnable 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   3270
         Width           =   1380
      End
      Begin VB.CommandButton cmdExpand 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Enlarge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -64905
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   3660
         Width           =   945
      End
      Begin VB.TextBox txtLastSoldDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -72675
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   3225
         Width           =   1860
      End
      Begin VB.TextBox txtLastSoldQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -70785
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   3225
         Width           =   555
      End
      Begin VB.TextBox txtLastSoldPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -70185
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   3225
         Width           =   960
      End
      Begin VB.TextBox txtProductType 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   -73635
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   450
         Width           =   2745
      End
      Begin VB.TextBox txtDefaultDeliveryDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   -65490
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   3540
         Width           =   1395
      End
      Begin VB.TextBox txtSS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   -73140
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   3525
         Width           =   1395
      End
      Begin VB.TextBox txtLastCountedPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -70170
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   2805
         Width           =   960
      End
      Begin VB.TextBox txtLastCountedQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -70785
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   2805
         Width           =   555
      End
      Begin VB.TextBox txtLastCounted 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -72675
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   2805
         Width           =   1860
      End
      Begin VB.TextBox txtPublisher 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1210
         Width           =   4980
      End
      Begin VB.TextBox txtEdition 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1600
         Width           =   4980
      End
      Begin VB.TextBox txtPubDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   820
         Width           =   4980
      End
      Begin VB.TextBox txtPubPlace 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   430
         Width           =   4980
      End
      Begin VB.TextBox txtLastReceived 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -72690
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   2385
         Width           =   1875
      End
      Begin VB.TextBox txtLastReceivedQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -70785
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   2385
         Width           =   555
      End
      Begin VB.TextBox txtLastReceivedPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -70170
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   2385
         Width           =   960
      End
      Begin VB.TextBox txtLastOrderedPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -70170
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   1965
         Width           =   960
      End
      Begin VB.TextBox txtLastOrderedQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -70785
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1965
         Width           =   555
      End
      Begin VB.TextBox txtLastOrdered 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -72690
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1965
         Width           =   1875
      End
      Begin VB.TextBox txtDateLastModified 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -66360
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3540
         Width           =   2550
      End
      Begin VB.TextBox txtDateAdded 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -66360
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   3120
         Width           =   2550
      End
      Begin VB.Frame Frame2 
         Caption         =   "Aged stock"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   915
         Left            =   -74355
         TabIndex        =   38
         Top             =   480
         Width           =   5220
         Begin VB.TextBox txt18Plus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   2415
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txtAgedDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   375
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   330
            Width           =   1170
         End
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -74715
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   1590
         Width           =   7230
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -74715
         MultiLine       =   -1  'True
         TabIndex        =   34
         Top             =   690
         Width           =   7230
      End
      Begin VB.TextBox txtUSPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2905
         Width           =   1380
      End
      Begin VB.TextBox txtUKPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2545
         Width           =   1380
      End
      Begin VB.TextBox txtTotalSold 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   405
         Left            =   2130
         TabIndex        =   23
         Top             =   825
         Width           =   750
      End
      Begin VB.TextBox txtCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   1245
         TabIndex        =   22
         Top             =   2205
         Width           =   1380
      End
      Begin VB.TextBox txtSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   1245
         TabIndex        =   21
         Top             =   1845
         Width           =   1380
      End
      Begin VB.TextBox txtRRP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   1245
         TabIndex        =   20
         Top             =   1485
         Width           =   1380
      End
      Begin VB.TextBox txtReserved 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   405
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   825
         Width           =   750
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   405
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   825
         Width           =   750
      End
      Begin VB.TextBox txtSection 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   645
         Left            =   -73635
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   870
         Width           =   2760
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   -65490
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   1395
      End
      Begin VB.TextBox txtBinding 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
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
         Height          =   330
         Left            =   -65490
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3150
         Width           =   1395
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   2790
         Left            =   -74835
         OleObjectBlob   =   "frmProductPrev.frx":0767
         TabIndex        =   17
         Top             =   460
         Width           =   10965
      End
      Begin MSChart20Lib.MSChart chart1 
         Height          =   2415
         Left            =   -74805
         OleObjectBlob   =   "frmProductPrev.frx":49CD
         TabIndex        =   72
         Top             =   315
         Visible         =   0   'False
         Width           =   10785
      End
      Begin MSChart20Lib.MSChart Chart2 
         Height          =   870
         Left            =   -74640
         OleObjectBlob   =   "frmProductPrev.frx":11874
         TabIndex        =   91
         Top             =   2670
         Width           =   10425
      End
      Begin TrueOleDBGrid60.TDBGrid StGrid 
         Height          =   3495
         Left            =   3000
         OleObjectBlob   =   "frmProductPrev.frx":638EF
         TabIndex        =   104
         Top             =   375
         Width           =   10650
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Catalogue heading"
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
         Height          =   270
         Left            =   -74790
         TabIndex        =   99
         Top             =   3165
         Width           =   1680
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Flag text"
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
         Height          =   225
         Left            =   -74670
         TabIndex        =   97
         Top             =   2205
         Width           =   705
      End
      Begin VB.Label Label41 
         BackColor       =   &H00CECECE&
         Caption         =   "Estimate of whether item was in stock per week (current week on right)"
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
         Height          =   255
         Left            =   -72480
         TabIndex        =   94
         Top             =   3555
         Width           =   6510
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Returnable"
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
         Height          =   285
         Left            =   210
         TabIndex        =   93
         Top             =   3315
         Width           =   975
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "Last sold"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74220
         TabIndex        =   89
         Top             =   3255
         Width           =   1395
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   -70080
         TabIndex        =   85
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   -70935
         TabIndex        =   84
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   -72195
         TabIndex        =   83
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Product type"
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
         Height          =   285
         Left            =   -74865
         TabIndex        =   82
         Top             =   465
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "U.S. Price"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1245
         TabIndex        =   80
         Top             =   3660
         Width           =   1590
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Lead time"
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
         Height          =   285
         Left            =   -67110
         TabIndex        =   79
         Top             =   3570
         Width           =   1590
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Distributor"
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
         Height          =   240
         Left            =   -70740
         TabIndex        =   77
         Top             =   2055
         Width           =   1590
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Last deal"
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
         Height          =   240
         Left            =   -70335
         TabIndex        =   76
         Top             =   2445
         Width           =   1170
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Order by seesafe"
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
         Height          =   285
         Left            =   -74895
         TabIndex        =   74
         Top             =   3570
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Last counted"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74220
         TabIndex        =   71
         Top             =   2835
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher"
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
         Height          =   255
         Left            =   -70125
         TabIndex        =   67
         Top             =   1255
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Edition"
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
         Height          =   255
         Left            =   -69795
         TabIndex        =   66
         Top             =   1660
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Publication date"
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
         Height          =   255
         Left            =   -70620
         TabIndex        =   65
         Top             =   850
         Width           =   1470
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Publication place"
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
         Height          =   255
         Left            =   -70725
         TabIndex        =   64
         Top             =   445
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Last received"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74220
         TabIndex        =   59
         Top             =   2415
         Width           =   1395
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   -69090
         TabIndex        =   53
         Top             =   1995
         Width           =   3135
      End
      Begin VB.Label lblDeal 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -69090
         TabIndex        =   52
         Top             =   2385
         Width           =   2295
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Last ordered"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74220
         TabIndex        =   51
         Top             =   1995
         Width           =   1395
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Last modified"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -67905
         TabIndex        =   50
         Top             =   3585
         Width           =   1395
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Added"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -67815
         TabIndex        =   49
         Top             =   3195
         Width           =   1290
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Comment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74685
         TabIndex        =   37
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74715
         TabIndex        =   36
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "U.S. Price"
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
         Height          =   285
         Left            =   315
         TabIndex        =   33
         Top             =   2948
         Width           =   870
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "U.K. Price"
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
         Height          =   285
         Left            =   315
         TabIndex        =   32
         Top             =   2584
         Width           =   870
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Total sold"
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
         Height          =   285
         Left            =   2040
         TabIndex        =   31
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Cost"
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
         Height          =   285
         Left            =   435
         TabIndex        =   30
         Top             =   2235
         Width           =   750
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "S.P."
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
         Height          =   285
         Left            =   435
         TabIndex        =   29
         Top             =   1860
         Width           =   750
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "R.R.P."
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
         Height          =   285
         Left            =   435
         TabIndex        =   28
         Top             =   1500
         Width           =   750
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Reserved"
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
         Height          =   285
         Left            =   1050
         TabIndex        =   27
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "On hand"
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
         Height          =   285
         Left            =   225
         TabIndex        =   26
         Top             =   480
         Width           =   705
      End
      Begin VB.Label lblObsolete 
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -71415
         TabIndex        =   16
         Top             =   3525
         Width           =   1380
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Sections"
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
         Height          =   285
         Left            =   -74880
         TabIndex        =   11
         Top             =   885
         Width           =   1080
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "V.A.T. Rate"
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
         Height          =   285
         Left            =   -66435
         TabIndex        =   10
         Top             =   2790
         Width           =   915
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Binding"
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
         Height          =   285
         Left            =   -66600
         TabIndex        =   9
         Top             =   3180
         Width           =   1080
      End
   End
   Begin VB.TextBox txtSubtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   810
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   5760
   End
   Begin VB.TextBox txtAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   810
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   105
      Width           =   5760
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   810
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   525
      Width           =   5760
   End
   Begin VB.Label lblNonStock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      Caption         =   "NON-STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   -90
      TabIndex        =   75
      Top             =   6120
      Width           =   3705
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
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
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   150
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtitle"
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
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   975
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
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
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   555
      Width           =   495
   End
End
Attribute VB_Name = "frmProductPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngID As Long
Public WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim flgLoading As Boolean
Dim tlCatHead As z_TextList
Private tlSections As z_TextList
Private tlProductTypes As z_TextList
Dim mCancel As Boolean
Dim XA As XArrayDB  'Copies grid
Dim XB As XArrayDB  'Stock grid
Dim XC As XArrayDB  'OSPOs
Dim XD As XArrayDB  'OSCOs
Dim XE As XArrayDB  'OSAppros
Dim XF As XArrayDB  'Movements
Dim strTime As String
Public Property Get Timing() As String
    Timing = strTime
End Property
Sub Component(pProduct As a_Product, Optional pstrTime As String)
    strTime = pstrTime
strTime = strTime & "Start frmProductPrev:component:" & Now() & vbCrLf
    Set oProd = Nothing
    Set oProd = pProduct
strTime = strTime & "End frmProductPrev:component:" & Now() & vbCrLf
End Sub
Public Property Get CategoryName() As String
    CategoryName = tlSections.Item(oProd.CategoryID)
End Property
Public Property Get ProductTypeName() As String
    ProductTypeName = tlProductTypes.Item(oProd.ProductTypeID)
End Property
Private Sub cmdDelete_Click()

    If MsgBox("Confirm deletion of product: " & oProd.Title & vbCrLf & "This is permanent!", vbOKCancel + vbExclamation, "Confirm") = vbOK Then
        oProd.Delete
        oProd.ApplyEdit
        Unload Me
    End If
End Sub
Private Sub chart1_AxisLabelSelected(axisID As Integer, AxisIndex As Integer, labelSetIndex As Integer, LabelIndex As Integer, MouseFlags As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub chart1_AxisTitleSelected(axisID As Integer, AxisIndex As Integer, MouseFlags As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub chart1_ChartSelected(MouseFlags As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub chart1_FootnoteSelected(MouseFlags As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub chart1_LegendSelected(MouseFlags As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub chart1_PlotSelected(MouseFlags As Integer, Cancel As Integer)
Cancel = True
End Sub

Private Sub chart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
''MsgBox "Series " & Series
''MsgBox "Datapoint " & DataPoint
'    Select Case Series
'    Case 1
'        MsgBox "Sales qty = " & oProd.CurrentSales.FindByWeek(DataPoint).Qty & vbCrLf & "Sales value = " & oProd.CurrentSales.FindByWeek(DataPoint).ValuF
'    Case 2
'   '     MsgBox "Sales qty = " & oProd.PreviousSales.FindByWeek(DataPoint).Qty & vbCrLf & "Sales value = " & oProd.PreviousSales.FindByWeek(DataPoint).ValuF
'    End Select
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
Dim frm As frmProduct
    Set frm = New frmProduct
    frm.Component oProd, Me
    frm.Show
    Exit Sub
Errh:
    MsgBox Error
End Sub

Private Sub cmdsearchisbn_Click()
    Set oProd = Nothing
    Set oProd = New a_Product
    With oProd
    .Load "", 0, txtisbnsearch
       
    txtAuthor = .Author
    Me.Caption = "Stock code: " & .code
    txtSubtitle = .Subtitle
    txtTitle = .Title
    txtPublisher = .Publisher
        
    End With
    LoadControls
    LoadStock
End Sub




Private Sub Form_Resize()
Dim lngDiff As Long
On Error Resume Next
    SSTab1.Width = (Me.Width - 600)
    lngDiff = SSTab1.Height
    SSTab1.Height = Me.Height - 3100
    lngDiff = SSTab1.Height - lngDiff
    StGrid.Width = Me.Width - 4000
    StGrid.Height = StGrid.Height + lngDiff
    Me.cmdClose.top = cmdClose.top + lngDiff
    Me.cmdEdit.top = cmdEdit.top + lngDiff
    Me.cmdExpand.top = cmdExpand.top + lngDiff
    Me.cmdsearchisbn.top = cmdsearchisbn.top + lngDiff
    Me.Frame1.top = Frame1.top + lngDiff
    Me.lblNonStock.top = lblNonStock.top + lngDiff
    Me.Text1.top = Text1.top + lngDiff
    Me.chkExSales.top = chkExSales.top + lngDiff

End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
            Cancel = True
End Sub

Private Sub StGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Cancel = True
End Sub



Private Sub Form_Load()
    left = 10
    top = 10
    Width = 11700
    Height = 6800
    Set tlSections = New z_TextList
    Set tlProductTypes = New z_TextList
    tlSections.Load ltDictionary, , dtCategory
    tlProductTypes.Load ltProductType
    LoadControls
    Me.SSTab1.Tab = 0
End Sub
Public Sub RefreshForm()
    LoadControls
End Sub
Private Sub LoadControls()
    flgLoading = True
    Me.Caption = "Stock code: " & oProd.code & "      EAN: " & oProd.EAN
    txtTitle = oProd.Title
    txtSubtitle = oProd.Subtitle
    txtAuthor = oProd.Author
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    txtPubPlace = oProd.PublicationPlace
    Me.txtPubDate = oProd.PublicationDate
    Me.txtPubDate = oProd.PublicationDate
    Me.txtBinding = oProd.BindingCode
    Me.txtSection = tlProductTypes.Item(oProd.CategoryID)
    Me.txtProductType = tlProductTypes.Item(oProd.ProductTypeID)
    txtDefaultDeliveryDays = oProd.DefaultDeliveryDays
    Me.chkExSales = IIf(oProd.ExcludeFromSales, 1, 0)
    Me.txtRRP = oProd.RRPF
    Me.txtSP = oProd.SPF
    Me.txtCost = oProd.CostF
    Me.txtTotalSold = oProd.QtyTotalSold
    Me.txtUKPrice = oProd.UKPriceF
    Me.txtUSPrice = oProd.USPriceF
    Me.txtAgedDate = ""
    Me.txtComment = oProd.Comment
    Me.txtCost = oProd.CostF
    Me.txtDateAdded = oProd.DateRecordAddedF
    Me.txtDateLastModified = oProd.DateLastModifiedF
    Me.txtDescription = oProd.Description
    Me.txtLastCounted = oProd.dateLastCountedF
    Me.txtLastCountedPrice = oProd.PriceLastCountedF
    Me.txtLastCountedQty = oProd.QtyLastCountedF
    Me.txtLastReceived = oProd.DateLastDeliveredF
    Me.txtLastReceivedPrice = oProd.PriceLastDeliveredF
    Me.txtLastReceivedQty = oProd.QtyLastDeliveredF
    Me.txtLastOrdered = oProd.DateLastOrderedF
    Me.txtLastOrderedPrice = oProd.PriceLastOrderedF
    Me.txtLastOrderedQty = oProd.QtyLastOrderedF
    Me.txtLastSoldDate = oProd.DateLastSoldF
    Me.txtLastSoldPrice = oProd.PriceLastSoldF
    Me.txtLastSoldQty = oProd.QtylastSold
    Me.txtOnHand = oProd.QtyOnHandF
'    Me.txtSummary = oProd.Summary
    Me.txtVAT = oProd.VATRateToUseF
    Me.txtReserved = oProd.QtyReservedF
    Me.txtReturnable = oProd.ReturnAvailability
    txtFlagText = oProd.FlagText
    txtBIC = oProd.BIC
    txtBICDescription = oPC.Configuration.BICs.FetchBICDescriptionsFromCodeSet(txtBIC)
'    txtBICDescription = oProd.BICDescription
    Me.lblNonStock.Visible = oProd.isNonStock
    Me.lblObsolete = IIf(oProd.Obsolete, "obsolete", "")
    txtSS = IIf(oProd.Seesafe = 1, "Yes", "")
    Me.lblSupplier.Caption = oProd.lastsuppliername
    Me.lblDeal.Caption = oProd.lastDealDescription
    lblStatus = oProd.statusF
   ' DTMMSince.Value = IIf(oProd.DateLastCounted < DTMMSince.MinDate, DTMMSince.MinDate, oProd.DateLastCounted)
    LoadStock
    flgLoading = False
End Sub
Private Sub LoadStock()
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XB = New XArrayDB
    XB.Clear
    XB.ReDim 1, oProd.Stores.Count, 1, 10
    For lngIndex = 1 To oProd.Stores.Count
        XB.Value(lngIndex, 1) = oProd.Stores(lngIndex).StoreName
        XB.Value(lngIndex, 2) = oProd.Stores(lngIndex).QtyOnHand
        XB.Value(lngIndex, 3) = oProd.Stores(lngIndex).QtyonOrder
        XB.Value(lngIndex, 4) = oProd.Stores(lngIndex).QtyOnBackorder
        XB.Value(lngIndex, 5) = oProd.Stores(lngIndex).SP
        XB.Value(lngIndex, 6) = oProd.Stores(lngIndex).TotalQtySold
        XB.Value(lngIndex, 7) = oProd.Stores(lngIndex).LastSoldDateF
        XB.Value(lngIndex, 8) = oProd.Stores(lngIndex).LastReceivedF
        XB.Value(lngIndex, 9) = oProd.Stores(lngIndex).LastDeliveredPriceF
        XB.Value(lngIndex, 10) = oProd.Stores(lngIndex).LastOrderedDateF2
       ' XB.Value(lngIndex, 11) = oProd.Stores(lngIndex).QtyLastStocktake
    Next
    XB.QuickSort 1, oProd.Stores.Count, 1, XORDER_ASCEND, XTYPE_STRING
    StGrid.Array = XB
    StGrid.ReBind

End Sub

