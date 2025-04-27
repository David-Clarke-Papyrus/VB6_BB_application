VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmProductPrev 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Stock"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11820
   ControlBox      =   0   'False
   Icon            =   "frmProductPrev.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleMode       =   0  'User
   ScaleWidth      =   15592.34
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Sales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   124
      Top             =   7440
      Width           =   1200
   End
   Begin VB.CommandButton cmdSales 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Sales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10275
      Style           =   1  'Graphical
      TabIndex        =   111
      Top             =   1335
      Width           =   1200
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
      Left            =   10545
      Picture         =   "frmProductPrev.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5445
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
      Left            =   9450
      Picture         =   "frmProductPrev.frx":03B5
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5445
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
      Left            =   240
      TabIndex        =   6
      Top             =   1395
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   535
      BackColor       =   13882315
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
      Tab(0).Control(10)=   "Label29"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblNDA"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "MMGRID"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "APPGRID"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "COGrid"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "POGrid"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtOnHand"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtReserved"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtRRP"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtSP"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCost"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtTotalSold"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtUKPrice"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtUSPrice"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtTotalOSPO"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtTotalOSCO"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtTotalOSAP"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cmdALLMM"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "DTMMSince"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtReturnable"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtSSP"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdRecon"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProductPrev.frx":06DB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbPSECs"
      Tab(1).Control(1)=   "txtLoyaltyRate"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "txtProductType"
      Tab(1).Control(4)=   "txtDefaultDeliveryDays"
      Tab(1).Control(5)=   "txtSS"
      Tab(1).Control(6)=   "txtPublisher"
      Tab(1).Control(7)=   "txtEdition"
      Tab(1).Control(8)=   "txtPubDate"
      Tab(1).Control(9)=   "txtPubPlace"
      Tab(1).Control(10)=   "txtVAT"
      Tab(1).Control(11)=   "txtBinding"
      Tab(1).Control(12)=   "Label11"
      Tab(1).Control(13)=   "Label35"
      Tab(1).Control(14)=   "Label34"
      Tab(1).Control(15)=   "Label33"
      Tab(1).Control(16)=   "Label27"
      Tab(1).Control(17)=   "Label23"
      Tab(1).Control(18)=   "Label7"
      Tab(1).Control(19)=   "Label8"
      Tab(1).Control(20)=   "Label6"
      Tab(1).Control(21)=   "Label9"
      Tab(1).Control(22)=   "lblSupplier"
      Tab(1).Control(23)=   "lblDeal"
      Tab(1).Control(24)=   "Label26"
      Tab(1).Control(25)=   "Label10"
      Tab(1).Control(26)=   "Label20"
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "&3. Notes"
      TabPicture(2)   =   "frmProductPrev.frx":06F7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCatalogues"
      Tab(2).Control(1)=   "txtCategoryHeading"
      Tab(2).Control(2)=   "txtFlagText"
      Tab(2).Control(3)=   "txtComment"
      Tab(2).Control(4)=   "txtDescription"
      Tab(2).Control(5)=   "Label25"
      Tab(2).Control(6)=   "Label17"
      Tab(2).Control(7)=   "Label2"
      Tab(2).Control(8)=   "Label30"
      Tab(2).Control(9)=   "Label28"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "&4. Statistics"
      TabPicture(3)   =   "frmProductPrev.frx":0713
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtLastSoldDate"
      Tab(3).Control(1)=   "txtLastSoldQty"
      Tab(3).Control(2)=   "txtLastSoldPrice"
      Tab(3).Control(3)=   "txtLastCountedPrice"
      Tab(3).Control(4)=   "txtLastCountedQty"
      Tab(3).Control(5)=   "txtLastCounted"
      Tab(3).Control(6)=   "txtLastReceived"
      Tab(3).Control(7)=   "txtLastReceivedQty"
      Tab(3).Control(8)=   "txtLastReceivedPrice"
      Tab(3).Control(9)=   "txtLastOrderedPrice"
      Tab(3).Control(10)=   "txtLastOrderedQty"
      Tab(3).Control(11)=   "txtLastOrdered"
      Tab(3).Control(12)=   "txtDateLastModified"
      Tab(3).Control(13)=   "txtDateAdded"
      Tab(3).Control(14)=   "Frame2"
      Tab(3).Control(15)=   "Label42"
      Tab(3).Control(16)=   "Label41"
      Tab(3).Control(17)=   "Label39"
      Tab(3).Control(18)=   "Label38"
      Tab(3).Control(19)=   "Label37"
      Tab(3).Control(20)=   "Label36"
      Tab(3).Control(21)=   "Label1"
      Tab(3).Control(22)=   "Label15"
      Tab(3).Control(23)=   "Label24"
      Tab(3).Control(24)=   "Label22"
      Tab(3).Control(25)=   "Label12"
      Tab(3).ControlCount=   26
      TabCaption(4)   =   "&5. Substitute products"
      TabPicture(4)   =   "frmProductPrev.frx":072F
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label44"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label43"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "GSF"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "GSB"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "cmdSubs"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "cmdSub"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "&6. Copies"
      TabPicture(5)   =   "frmProductPrev.frx":074B
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grid1"
      Tab(5).ControlCount=   1
      Begin VB.CommandButton cmdRecon 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Recon."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7410
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   3510
         Width           =   645
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
         Left            =   -72690
         Locked          =   -1  'True
         TabIndex        =   122
         Top             =   3225
         Width           =   1860
      End
      Begin VB.CommandButton cmdSub 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Create subst."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -65220
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   540
         Width           =   1305
      End
      Begin VB.CommandButton cmdSubs 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Fetch substitutions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -74820
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   450
         Width           =   1920
      End
      Begin VB.TextBox txtSSP 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   1905
         Width           =   1200
      End
      Begin VB.TextBox txtCatalogues 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   1395
         Left            =   -67170
         MultiLine       =   -1  'True
         TabIndex        =   108
         Top             =   690
         Width           =   2970
      End
      Begin VB.ListBox lbPSECs 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   810
         Left            =   -73620
         TabIndex        =   106
         Top             =   885
         Width           =   2715
      End
      Begin VB.TextBox txtLoyaltyRate 
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
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   2775
         Width           =   1395
      End
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
         TabIndex        =   102
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
            TabIndex        =   103
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
         TabIndex        =   99
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
         Left            =   -74730
         Locked          =   -1  'True
         TabIndex        =   97
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
         TabIndex        =   95
         Top             =   3270
         Width           =   1380
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
         TabIndex        =   93
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
         TabIndex        =   92
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
         TabIndex        =   87
         Top             =   450
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker DTMMSince 
         Height          =   390
         Left            =   9825
         TabIndex        =   85
         Top             =   3510
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   16056321
         CurrentDate     =   37656
         MaxDate         =   55153
         MinDate         =   34820
      End
      Begin VB.CommandButton cmdALLMM 
         BackColor       =   &H00C4BCA4&
         Caption         =   " All movements since:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8085
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3510
         Width           =   1725
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
         TabIndex        =   82
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
         TabIndex        =   77
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
         Left            =   -70185
         Locked          =   -1  'True
         TabIndex        =   74
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
         TabIndex        =   73
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
         Left            =   -72690
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   2805
         Width           =   1860
      End
      Begin VB.TextBox txtTotalOSAP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
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
         Left            =   7260
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   1645
         Width           =   3705
      End
      Begin VB.TextBox txtTotalOSCO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
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
         Left            =   3435
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   3510
         Width           =   3705
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
         Top             =   430
         Width           =   4980
      End
      Begin VB.TextBox txtTotalOSPO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   1650
         Width           =   3705
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
         Left            =   -72705
         Locked          =   -1  'True
         TabIndex        =   55
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
         TabIndex        =   54
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
         Left            =   -70185
         Locked          =   -1  'True
         TabIndex        =   110
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
         Left            =   -70185
         Locked          =   -1  'True
         TabIndex        =   45
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
         TabIndex        =   44
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
         Left            =   -72705
         Locked          =   -1  'True
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   35
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
         TabIndex        =   32
         Top             =   1590
         Width           =   7230
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -74715
         MultiLine       =   -1  'True
         TabIndex        =   31
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   645
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
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1905
         Width           =   1170
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1305
         Width           =   1200
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
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1305
         Width           =   1170
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
         TabIndex        =   116
         Top             =   645
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
         TabIndex        =   15
         Top             =   645
         Width           =   750
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
      Begin TrueOleDBGrid60.TDBGrid POGrid 
         Height          =   1230
         Left            =   3075
         OleObjectBlob   =   "frmProductPrev.frx":0767
         TabIndex        =   57
         Top             =   420
         Width           =   4200
      End
      Begin TrueOleDBGrid60.TDBGrid COGrid 
         Height          =   1335
         Left            =   3090
         OleObjectBlob   =   "frmProductPrev.frx":50BA
         TabIndex        =   68
         Top             =   2145
         Width           =   4185
      End
      Begin TrueOleDBGrid60.TDBGrid APPGRID 
         Height          =   1215
         Left            =   7410
         OleObjectBlob   =   "frmProductPrev.frx":9555
         TabIndex        =   70
         Top             =   435
         Width           =   3885
      End
      Begin TrueOleDBGrid60.TDBGrid MMGRID 
         Height          =   1335
         Left            =   7395
         OleObjectBlob   =   "frmProductPrev.frx":D9F1
         TabIndex        =   76
         Top             =   2130
         Width           =   3900
      End
      Begin TrueOleDBGrid60.TDBGrid GSB 
         Height          =   1320
         Left            =   -73410
         OleObjectBlob   =   "frmProductPrev.frx":11AF4
         TabIndex        =   118
         Top             =   2550
         Width           =   9495
      End
      Begin TrueOleDBGrid60.TDBGrid GSF 
         Height          =   1500
         Left            =   -73410
         OleObjectBlob   =   "frmProductPrev.frx":15AF0
         TabIndex        =   119
         Top             =   960
         Width           =   9495
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   2790
         Left            =   -74790
         OleObjectBlob   =   "frmProductPrev.frx":19C1C
         TabIndex        =   123
         Top             =   660
         Width           =   10965
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Possible substitutions for this item"
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   -74760
         TabIndex        =   121
         Top             =   2970
         Width           =   1185
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "This item can substitute for these . . ."
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   -74760
         TabIndex        =   120
         Top             =   1500
         Width           =   1185
      End
      Begin VB.Label Label42 
         BackStyle       =   0  'Transparent
         Caption         =   "(Cost)"
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
         Left            =   -69150
         TabIndex        =   115
         Top             =   2430
         Width           =   765
      End
      Begin VB.Label Label41 
         BackStyle       =   0  'Transparent
         Caption         =   "(R.R.P.)"
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
         Left            =   -69150
         TabIndex        =   114
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label lblNDA 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No discount allowed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   180
         TabIndex        =   113
         Top             =   2250
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Special"
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
         Left            =   1620
         TabIndex        =   112
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Catalogues"
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
         Left            =   -67185
         TabIndex        =   109
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loyalty Rate"
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
         Left            =   -70035
         TabIndex        =   105
         Top             =   2805
         Width           =   915
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   100
         Top             =   3165
         Width           =   1680
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   98
         Top             =   2205
         Width           =   705
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   96
         Top             =   3315
         Width           =   975
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   94
         Top             =   3255
         Width           =   1395
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         TabIndex        =   91
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         TabIndex        =   90
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         TabIndex        =   89
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   88
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
         TabIndex        =   86
         Top             =   3660
         Width           =   1590
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   83
         Top             =   3570
         Width           =   1590
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   81
         Top             =   2055
         Width           =   1590
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   80
         Top             =   2445
         Width           =   1170
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   78
         Top             =   3570
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   75
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
         TabIndex        =   66
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
         TabIndex        =   117
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
         TabIndex        =   64
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
         TabIndex        =   63
         Top             =   445
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   -74940
         TabIndex        =   56
         Top             =   2415
         Width           =   2115
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
         TabIndex        =   50
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
         TabIndex        =   49
         Top             =   2385
         Width           =   2295
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   -74880
         TabIndex        =   48
         Top             =   1995
         Width           =   2055
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   47
         Top             =   3585
         Width           =   1395
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   46
         Top             =   3195
         Width           =   1290
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   34
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   33
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   30
         Top             =   2948
         Width           =   870
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   29
         Top             =   2584
         Width           =   870
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   28
         Top             =   390
         Width           =   870
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cost (avg.)"
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
         Left            =   105
         TabIndex        =   27
         Top             =   1665
         Width           =   1110
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sell.P."
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
         Left            =   1695
         TabIndex        =   26
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   225
         TabIndex        =   25
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   24
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         TabIndex        =   23
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         BackStyle       =   0  'Transparent
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
         BackStyle       =   0  'Transparent
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
      Width           =   4455
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
   Begin TrueOleDBGrid60.TDBGrid StGrid 
      Height          =   1200
      Left            =   6645
      OleObjectBlob   =   "frmProductPrev.frx":1DE82
      TabIndex        =   71
      Top             =   90
      Width           =   4815
   End
   Begin VB.Label lblObsolete 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5775
      TabIndex        =   107
      Top             =   5640
      Width           =   2940
   End
   Begin VB.Label lblNonStock 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "NON-STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   3525
      TabIndex        =   79
      Top             =   5670
      Width           =   1785
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
'Private lslist As ListItem


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
Dim XSF As XArrayDB 'substitutes for these
Dim XSB As XArrayDB 'can be substituted by
Dim strTime As String
Public Property Get Timing() As String
    Timing = strTime
End Property
Sub Component(pProduct As a_Product, Optional pstrTime As String)
    On Error GoTo errHandler
    strTime = pstrTime
    strTime = strTime & "Start frmProductPrev:component:" & Now() & vbCrLf
    Set oProd = Nothing
    Set oProd = pProduct

    Set tlCatHead = Nothing
    Set tlCatHead = New z_TextList
    tlCatHead.Load ltCatalogueHeadings



strTime = strTime & "End frmProductPrev:component:" & Now() & vbCrLf
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Component(pProduct,pstrTime)", Array(pProduct, pstrTime)
End Sub
Public Property Get CategoryName() As String
    CategoryName = tlSections.Item(oProd.CategoryID)
End Property
Public Property Get ProductTypeName() As String
    ProductTypeName = tlProductTypes.Item(oProd.ProductTypeID)
End Property



Private Sub cmdRecon_Click()
Dim frmMM As frmMovements
    oProd.ReloadRecentMovements
    Set frmMM = New frmMovements
    frmMM.Component oProd
    frmMM.Show
End Sub

Private Sub cmdSales_Click()
Dim frm As frmSalesCH
    Set frm = New frmSalesCH
    frm.Component oProd
    frm.Show
End Sub

Private Sub cmdSub_Click()
Dim frm As New frmSubstitute
    frm.Component2 Trim(oProd.pID)
    frm.Show
End Sub

Private Sub cmdSubs_Click()
    Set XSF = New XArrayDB
    Set XSB = New XArrayDB
    oProd.GetSubstitutes XSF, XSB
    Set GSF.Array = XSF
    Set GSB.Array = XSB
    GSF.ReBind
    GSB.ReBind
End Sub

Private Sub Command1_Click()
  '  ErrorIn "TEST ERROR CRASH for CLEANUP"
  oProd.crash
End Sub

Private Sub Form_Deactivate()
    UnsetMenu
End Sub
Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Form_Activate"
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuAdjust.Enabled = True
'    Forms(0).mnuProductPreview.Visible = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.SetMenu"
End Sub

Private Sub cboCatHead_Click()
'    oProd.setCatalogueheadingID tlCatHead.Key(cboCatHead)
End Sub

Private Sub cmdDelete_Click()

    If MsgBox("Confirm deletion of product: " & oProd.Title & vbCrLf & "This is permanent!", vbOKCancel + vbExclamation, "Confirm") = vbOK Then
        oProd.Delete
        oProd.ApplyEdit
        Unload Me
    End If


End Sub


Private Sub APPGRID_DblClick()
Dim frm As New frmAPPPreview
    If IsNull(APPGRID.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    frm.Component XE(APPGRID.Bookmark, 6)
    frm.Show
    Screen.MousePointer = vbDefault
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
'MsgBox "Series " & Series
'MsgBox "Datapoint " & DataPoint
    Select Case Series
    Case 1
        MsgBox "Sales qty = " & oProd.CurrentSales.FindByWeek(DataPoint).Qty & vbCrLf & "Sales value = " & oProd.CurrentSales.FindByWeek(DataPoint).ValuF
    Case 2
   '     MsgBox "Sales qty = " & oProd.PreviousSales.FindByWeek(DataPoint).Qty & vbCrLf & "Sales value = " & oProd.PreviousSales.FindByWeek(DataPoint).ValuF
    End Select
End Sub

Private Sub chart1_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
'Cancel = True
End Sub

Private Sub cmdALLMM_Click()
Dim frmMM As frmMovements_All
    oProd.ReloadMovements Me.DTMMSince
    Set frmMM = New frmMovements_All
    frmMM.Component oProd
    frmMM.Show
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



Private Sub formatgrid()
End Sub
Private Sub cmdsearchisbn_Click()
Dim lngResult As Long

    If Trim(txtisbnsearch) = "" Then Exit Sub
    Set oProd = Nothing
    Set oProd = New a_Product
    With oProd
        lngResult = .Load("", 0, txtisbnsearch)
        If lngResult = 99 Then
            MsgBox "Not found", vbInformation, "Status"
            Set oProd = Nothing
            Exit Sub
        End If
        txtAuthor = .Author
        Me.Caption = "Stock code: " & .code
        txtSubtitle = .SubTitle
        txtTitle = .Title
        txtPublisher = .Publisher
        
    End With
    LoadControls
    LoadCopies
    LoadStock
    LoadMovements
End Sub




Private Sub COGrid_DblClick()
Dim frm As New frmCOPreview
    If IsNull(COGrid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    frm.Component XD(COGrid.Bookmark, 6)
    frm.Show
    Screen.MousePointer = vbDefault
End Sub


Private Sub Grid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'Cancel = True
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
            Cancel = True
End Sub


Private Sub GSB_Click()
Dim str As String
    If IsNull(GSB.Bookmark) Then Exit Sub
    str = FNS(XSB.Value(GSB.Bookmark, 1))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)
End Sub

Private Sub GSF_Click()
Dim str As String
    If IsNull(GSF.Bookmark) Then Exit Sub
    str = FNS(XSF.Value(GSF.Bookmark, 1))
    If str = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText left(str, ISBNLENGTH)

End Sub
Private Sub GSF_DblClick()
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(GSF.Bookmark) Then Exit Sub
    str = FNS(XSF.Value(GSF.Bookmark, 1))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load "", 0, str, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.Component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.Component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub GSB_DblClick()
Dim frm As frmProductPrev
Dim frmA As frmProductPrevAQ
Dim oP As a_Product
Dim str As String
    If IsNull(GSB.Bookmark) Then Exit Sub
    str = FNS(XSB.Value(GSB.Bookmark, 1))
    If str = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    Set oP = New a_Product
    oP.Load "", 0, str, 0
    If oPC.Configuration.AntiquarianYN Then
        Set frmA = New frmProductPrevAQ
        frmA.Component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.Component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub MMGRID_DblClick()
Dim strType As String
Dim frm As Form
Dim i As Integer


    If IsNull(MMGRID.Bookmark) Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    strType = Trim(XF(MMGRID.Bookmark, 4))
    Select Case strType
    Case "APP"
        Set frm = New frmAPPPreview
        frm.Component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "APR"
        Set frm = New frmAPPRPreview
        frm.Component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "INV"
            Set frm = New frmInvoicePreview
            frm.Component XF(MMGRID.Bookmark, 6)
            frm.Show
    Case "CS", "POS"
        Set frm = New frmCSPreview
        frm.Component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "TF"
        Set frm = New frmTFPreview
        frm.Component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "DEL"
        Set frm = New frmDELPreview
        frm.Component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "CN"
        Set frm = New frmCNPreview
        frm.Component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "ADJ"
        MsgBox XF(MMGRID.Bookmark, 1), vbInformation, "ad-hoc adjustment"
    End Select
    Screen.MousePointer = vbDefault

End Sub

Public Sub mnuAdjust()
Dim frm As New frmStockAdjust
    frm.Component oProd
    frm.Show vbModal
    If frm.Cancelled = False Then
        
        Me.txtLastCounted = Now
        Me.txtLastCountedQty = frm.Counted
        Me.txtOnHand = frm.Count
    End If
    Unload frm
    Unload Me
End Sub

Private Sub POGrid_DblClick()
Dim frm As New frmPOPreview
    If IsNull(POGrid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    frm.Component XC(POGrid.Bookmark, 6)
    frm.Show
    Screen.MousePointer = vbDefault
End Sub



Private Sub StGrid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
'Cancel = True
End Sub

Private Sub StGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Cancel = True
End Sub

Private Sub Grid1_DblClick()
'MsgBox "Selected row is : " & Grid1.Row + 1
Dim frm As frmCopyPreview
Dim oCopy As a_Copy
    If IsNull(Grid1.Bookmark) Then Exit Sub
    Set oCopy = oProd.Copies(XA(Grid1.Bookmark, 7))
    Set frm = New frmCopyPreview

    frm.Component oCopy, oProd
    frm.Show ' vbModal

End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If XA(Bookmark, 5) > "" Then
        RowStyle.BackColor = &HDCDBF2
    End If
    If XA(Bookmark, 8) = True Then
        RowStyle.BackColor = vbRed
    End If
End Sub


Private Sub Form_Load()
    left = 10
    top = 10
    Width = 11700
    Height = 6800
    Set tlSections = oPC.Configuration.Sections  'New z_TextList
    Set tlProductTypes = oPC.Configuration.ProductTypes   'New z_TextList
  '  tlSections.Load ltDictionary, , dtCategory
  '  tlProductTypes.Load ltProductTypeAll
    LoadControls
    SSTab1.TabCaption(5) = "&6.Copies (" & oProd.Copies.CountForSale & ")"
    Me.SSTab1.Tab = 0
End Sub
Private Sub LoadProductSections()
Dim oPSEC As a_ProductSection
    With Me.lbPSECs
        .Clear
        For Each oPSEC In oProd.ProductSections
            .AddItem oPSEC.Description & "  " & oPSEC.PriorityF
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

End Sub

Public Sub RefreshForm()
    LoadControls
End Sub
Private Sub LoadControls()
    flgLoading = True
    Me.Caption = "Stock code: " & oProd.code & "      EAN: " & oProd.Ean
  '  Me.txtSection = oProd.Section
    Me.txtProductType = tlProductTypes.Item(oProd.ProductTypeID)
    txtTitle = oProd.Title
    txtSubtitle = oProd.SubTitle
    txtAuthor = oProd.Author
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    txtPubPlace = oProd.PublicationPlace
    Me.txtPubDate = oProd.PublicationDate
    Me.txtPubDate = oProd.PublicationDate
    Me.txtBinding = oProd.BindingCode
  '  Me.txtSection = tlProductTypes.Item(oProd.CategoryID)
    txtDefaultDeliveryDays = oProd.DefaultDeliveryDays
    Me.txtCategoryHeading = tlCatHead.Item(oProd.CatalogueheadingID)
    Me.txtRRP = oProd.RRPF
    If oProd.SpecialPrice > 0 Then
        txtSP.BackColor = &HDBFAFB
        txtSP.FontSize = 9
        txtSP.FontBold = False
        txtSP.ForeColor = vbGrayText
        txtSSP.BackColor = vbYellow
        txtSSP.FontSize = 10
        txtSSP.FontBold = True
        txtSSP.ForeColor = &H8000000D
        txtSP = "(" & oProd.SPF & ")"
        txtSSP = oProd.SpecialPriceF
    Else
        txtSP.BackColor = vbYellow
        txtSP.FontSize = 10
        txtSP.FontBold = True
        txtSP.ForeColor = &H8000000D
        txtSSP.BackColor = &HDBFAFB
        txtSSP.FontSize = 9
        txtSSP.FontBold = False
        txtSSP.ForeColor = vbGrayText
        txtSP = oProd.SPF
        txtSSP = oProd.SpecialPriceF
    End If
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
    Me.txtCatalogues = oProd.CatalogueEntries_Concat
    lblNDA.Visible = oProd.IsNDA
'    Me.txtSummary = oProd.Summary
    Me.txtVAT = oProd.VATRateToUseF
    Me.txtReserved = oProd.QtyReservedF
    Me.txtReturnable = oProd.ReturnAvailability
    txtFlagText = oProd.FlagText
    txtBIC = oProd.BIC
    txtLoyaltyRate = oProd.loyaltyRateF
    txtBICDescription = oPC.Configuration.BICs.FetchBICDescriptionsFromCodeSet(txtBIC)
'    txtBICDescription = oProd.BICDescription
    Me.lblNonStock.Visible = oProd.IsNONStock
    Me.lblObsolete = IIf(oProd.Obsolete, "Obsolete", "")
    txtSS = IIf(oProd.Seesafe = 1, "Yes", "")
    Me.lblSupplier.Caption = oProd.lastsuppliername
    Me.lblDeal.Caption = oProd.lastDealDescription
    lblStatus = oProd.statusF
    DTMMSince.Value = IIf(oProd.DateLastCounted < DTMMSince.MinDate, DTMMSince.MinDate, DateAdd("yyyy", -1, oProd.DateLastCounted))
    LoadMovements
    LoadCopies
    LoadStock
    LoadProductSections
    flgLoading = False
End Sub
Private Sub LoadCopies()
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String
Dim strCatalogues As String

   ' XA.Clear
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, oProd.Copies.Count, 1, 8
    For lngIndex = 1 To oProd.Copies.Count
        strCatalogues = ""
        For j = 1 To oProd.Copies(lngIndex).CatalogueEntries.Count
            strCatalogues = strCatalogues & oProd.Copies(lngIndex).CatalogueEntries(j).Serial
            If j < oProd.Copies(lngIndex).CatalogueEntries.Count Then strCatalogues = strCatalogues & ", "
        Next j
        XA.Value(lngIndex, 1) = oProd.Copies(lngIndex).Serial
        XA.Value(lngIndex, 2) = oProd.Copies(lngIndex).Description
        XA.Value(lngIndex, 3) = oProd.Copies(lngIndex).Comment
        XA.Value(lngIndex, 4) = oProd.Copies(lngIndex).PurchaseDateF
        XA.Value(lngIndex, 5) = oProd.Copies(lngIndex).SoldDateF
        XA.Value(lngIndex, 6) = oProd.Copies(lngIndex).PriceF
        XA.Value(lngIndex, 7) = oProd.Copies(lngIndex).Key
        XA.Value(lngIndex, 8) = oProd.Copies(lngIndex).IsDeleted
    Next
    XA.QuickSort 1, oProd.Copies.Count, 4, XORDER_DESCEND, XTYPE_DATE
    Grid1.Array = XA
    Grid1.ReBind
 '   Grid1.ReBind
End Sub
Private Sub LoadStock()
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XB = New XArrayDB
    XB.Clear
    XB.ReDim 1, oProd.Stores.Count, 1, 6
    lngIndex = 0
    Do While lngIndex < oProd.Stores.Count
        lngIndex = lngIndex + 1
        XB.Value(lngIndex, 1) = oProd.Stores(lngIndex).StoreName & oProd.Stores(lngIndex).LastSharedDateFShortwithParentheses
        XB.Value(lngIndex, 2) = oProd.Stores(lngIndex).QtyOnHand
        XB.Value(lngIndex, 3) = oProd.Stores(lngIndex).QtyReserved
        XB.Value(lngIndex, 4) = oProd.Stores(lngIndex).QtyOnBackorder
        XB.Value(lngIndex, 5) = oProd.Stores(lngIndex).QtyOnOrder
        XB.Value(lngIndex, 6) = oProd.Stores(lngIndex).QtyCopiesOnHand
    Loop
    XB.QuickSort 1, oProd.Stores.Count, 1, XORDER_ASCEND, XTYPE_STRING
    StGrid.Array = XB
    StGrid.ReBind

End Sub
Private Sub lvwCopies_BeforeLabelEdit(Cancel As Integer)
Cancel = True
End Sub

Private Sub LoadMovements()
Dim lngQtyOutstanding As Long
Dim lngQtySpecial As Long
Dim lngQtyApps As Long

Dim oOSPOD As d_OSSOrder
Dim oOSCOD As d_OSCOrder
Dim oOSAPP As d_OSAPPRO
Dim oMM As d_MM
Dim strMovementsSince As String

    oProd.LoadMovements
    LoadPOs
    LoadCOs
    LoadAPs
    LoadMMs
    lngQtyOutstanding = 0
    For Each oOSPOD In oProd.OSPOs
        lngQtyOutstanding = lngQtyOutstanding + oOSPOD.Firm + oOSPOD.SS - oOSPOD.QtyReceived
    Next
    lngQtySpecial = 0
    For Each oOSCOD In oProd.OSCOs
        lngQtySpecial = lngQtySpecial + oOSCOD.COLQty - oOSCOD.COLCollected
    Next
    lngQtyApps = 0
    For Each oOSAPP In oProd.OSAPs
        lngQtyApps = lngQtyApps + oOSAPP.APPQty - oOSAPP.APPReturned
    Next
    
    
    POGrid.Splits(0).Caption = "Purchase orders outstanding"
    COGrid.Splits(0).Caption = "Customer orders"
    APPGRID.Splits(0).Caption = "Appros issued"
    If oProd.DateLastCounted < "1995-01-01" Then
        MMGRID.Splits(0).Caption = "Mvmts since first received"
    Else
        MMGRID.Splits(0).Caption = "Mvmts since stocktake (" & oProd.dateLastCountedF & ":" & oProd.QtyLastCountedF & ")"
    End If
    
    txtTotalOSPO = "We are awaiting " & lngQtyOutstanding & " copies."
    txtTotalOSCO = "Customers are awaiting " & lngQtySpecial & " copies."
    txtTotalOSAP = "Expecting return of " & lngQtyApps & " copies."
End Sub
Private Sub LoadPOs()
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XC = New XArrayDB
    XC.Clear
    XC.ReDim 1, oProd.OSPOs.Count, 1, 7
    For lngIndex = 1 To oProd.OSPOs.Count
        XC.Value(lngIndex, 1) = oProd.OSPOs(lngIndex).DocCode
        XC.Value(lngIndex, 2) = oProd.OSPOs(lngIndex).DocDateF
        XC.Value(lngIndex, 3) = oProd.OSPOs(lngIndex).Firm
        XC.Value(lngIndex, 4) = oProd.OSPOs(lngIndex).SS
        XC.Value(lngIndex, 5) = oProd.OSPOs(lngIndex).QtyReceived
        XC.Value(lngIndex, 6) = oProd.OSPOs(lngIndex).TRID
        XC.Value(lngIndex, 7) = oProd.OSPOs(lngIndex).DateForSort
    Next
    XC.QuickSort 1, oProd.OSPOs.Count, 7, XORDER_DESCEND, XTYPE_STRING
    POGrid.Array = XC
    POGrid.ReBind

End Sub
Private Sub LoadCOs()
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XD = New XArrayDB
    XD.Clear
    XD.ReDim 1, oProd.OSCOs.Count, 1, 7
    For lngIndex = 1 To oProd.OSCOs.Count
        XD.Value(lngIndex, 1) = oProd.OSCOs(lngIndex).DocCode
        XD.Value(lngIndex, 2) = oProd.OSCOs(lngIndex).DocDateF
        XD.Value(lngIndex, 3) = oProd.OSCOs(lngIndex).TPName
        XD.Value(lngIndex, 4) = oProd.OSCOs(lngIndex).COLQty
        XD.Value(lngIndex, 5) = oProd.OSCOs(lngIndex).COLCollected
        XD.Value(lngIndex, 6) = oProd.OSCOs(lngIndex).TRID
        XD.Value(lngIndex, 7) = oProd.OSCOs(lngIndex).DateForSort
    Next
    XD.QuickSort 1, oProd.OSCOs.Count, 7, XORDER_DESCEND, XTYPE_STRING
    COGrid.Array = XD
    COGrid.ReBind

End Sub
Private Sub LoadAPs()
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XE = New XArrayDB
    XE.Clear
    XE.ReDim 1, oProd.OSAPs.Count, 1, 7
    For lngIndex = 1 To oProd.OSAPs.Count
        XE.Value(lngIndex, 1) = oProd.OSAPs(lngIndex).DocCode
        XE.Value(lngIndex, 2) = oProd.OSAPs(lngIndex).DocDateF
        XE.Value(lngIndex, 5) = oProd.OSAPs(lngIndex).TPName
        XE.Value(lngIndex, 3) = oProd.OSAPs(lngIndex).APPQty
        XE.Value(lngIndex, 4) = oProd.OSAPs(lngIndex).APPReturned
        XE.Value(lngIndex, 6) = oProd.OSAPs(lngIndex).TRID
        XE.Value(lngIndex, 7) = oProd.OSAPs(lngIndex).DocDate
    Next
    XE.QuickSort 1, oProd.OSAPs.Count, 7, XORDER_DESCEND, XTYPE_DATE
    APPGRID.Array = XE
    APPGRID.ReBind

End Sub
Private Sub LoadMMs()
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XF = New XArrayDB
    XF.Clear
    XF.ReDim 1, oProd.MMs.Count, 1, 7
    For lngIndex = 1 To oProd.MMs.Count
        XF.Value(lngIndex, 1) = oProd.MMs(lngIndex).DocCode
        XF.Value(lngIndex, 2) = oProd.MMs(lngIndex).DocDateF
        XF.Value(lngIndex, 3) = oProd.MMs(lngIndex).Qty
        XF.Value(lngIndex, 4) = oProd.MMs(lngIndex).Typ
        XF.Value(lngIndex, 5) = oProd.MMs(lngIndex).pID
        XF.Value(lngIndex, 6) = oProd.MMs(lngIndex).TRID
        XF.Value(lngIndex, 7) = oProd.MMs(lngIndex).Seq
    Next
    XF.QuickSort 1, oProd.MMs.Count, 7, XORDER_DESCEND, XTYPE_INTEGER
    MMGRID.Array = XF
    MMGRID.ReBind

End Sub

Private Sub Form_DblClick()
    If Not IsNull(oProd) Then
        Clipboard.Clear
        Clipboard.SetText oProd.productdetails
    End If
End Sub

Public Sub ExportInCatalogueFormat()
MsgBox "Hello"
'Dim oTF As New z_TextFile
'Dim strPath As String
'Dim strBillto As String
'Dim strDelto As String
'Dim strFOFile As String
'Dim strFilename As String
'Dim strXML As String
'Dim strCommand As String
'Dim i As Integer
'Dim strHTML As String
'Dim fs As New FileSystemObject
'Dim objXSL As New MSXML2.DOMDocument30
'Dim opXMLDOC As New MSXML2.DOMDocument30
'Dim objXMLDOC  As New MSXML2.DOMDocument30
'Dim strExecutable As String
'
'    Set xMLDoc = New ujXML
'    With xMLDoc
'        .docProgID = "MSXML2.DOMDocument"
'        .docInit "CatalogueExport_1"
'        .chCreate "CO"
'            .elText = "Customer orders at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
'        For i = 1 To mcol.Count
'
'            .elCreateSibling "DetailLine", True
'            .chCreate "Col_1"
'                .elText = mcol(i).TPName & (IIf(Len(Trim(mcol(i).TPACCNo)) <= 1, "", "(" & Trim(mcol(i).TPACCNo) & ")"))
'            .elCreateSibling "Col_2"
'                .elText = mcol(i).DocCode & mcol(i).StaffNameB
'            .elCreateSibling "Col_3"
'                .elText = mcol(i).DocDateF
'            .elCreateSibling "Col_4"
'                .elText = mcol(i).statusF
'                .navUP
'        Next i
'    End With
'
''FINALLY PRODUCE THE .XML FILE
'    strXML = oPC.SharedFolderRoot & "\TEMP\COs" & ".xml"
'    With xMLDoc
'        If fs.FileExists(strXML) Then
'            fs.DeleteFile strXML
'        End If
'        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
'    End With
'
'''WRITE THE .HTML FILE
'    objXSL.async = False
'    objXSL.validateOnParse = False
'    objXSL.resolveExternals = False
'    strPath = oPC.SharedFolderRoot & "\Templates\CO_RTF_1.xslt"
'    Set fs = New FileSystemObject
'    If fs.FileExists(strPath) Then
'        objXSL.Load strPath
'    End If
'
'    strFilename = oPC.LocalFolder & "\CO.RTF"
'    If fs.FileExists(strFilename) Then
'        fs.DeleteFile strFilename, True
'    End If
'    oTF.OpenTextFileToAppend strFilename
'    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
'    oTF.CloseTextFile
'
'    strExecutable = GetPDFExecutable(strFilename) & " " & strFilename
'    Shell strExecutable, vbNormalFocus
'
'    Exit Function

End Sub

