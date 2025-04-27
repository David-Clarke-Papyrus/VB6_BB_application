VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmProduct 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product master edit"
   ClientHeight    =   6645
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11430
   ControlBox      =   0   'False
   Icon            =   "frmProduct1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleMode       =   0  'User
   ScaleWidth      =   15077.87
   Begin VB.CheckBox chkCore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
      Caption         =   "Core"
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   8205
      TabIndex        =   81
      Top             =   135
      Width           =   780
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Cancel"
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
      Left            =   9285
      Picture         =   "frmProduct1.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   5865
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Left            =   10320
      Picture         =   "frmProduct1.frx":0396
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   5865
      Width           =   1000
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   5865
      Width           =   4350
   End
   Begin VB.CommandButton cmdChangeType 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Change this product type to a general product"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4575
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5910
      Width           =   3750
   End
   Begin VB.TextBox txtEAN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4770
      TabIndex        =   1
      Top             =   120
      Width           =   1620
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9075
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   750
      Width           =   2250
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3600
      Left            =   60
      TabIndex        =   10
      Top             =   2235
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   6350
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
      TabCaption(0)   =   "&1. Prices"
      TabPicture(0)   =   "frmProduct1.frx":0720
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdUnlockPrices"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtReason"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frMultibuys"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdUnlockMB"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frPrices"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProduct1.frx":073C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtWeight"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkMAG"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkObsolete"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cboProductType"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtLoyaltyRate"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdSupplier"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtVAT"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdSetDefault"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtBinding"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label23"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label11"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblSeesafe"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label40"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label15"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label14"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label2"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label10"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label20"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lblDeal"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "lblSupplier"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "&3. Notes && publishers status"
      TabPicture(2)   =   "frmProduct1.frx":0758
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtDescription"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtComment"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label28"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label30"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "&4. Wants"
      TabPicture(3)   =   "frmProduct1.frx":0774
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvwWants"
      Tab(3).Control(1)=   "cmdNewWant"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "&5. Catalogue"
      TabPicture(4)   =   "frmProduct1.frx":0790
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(1)=   "cboCatHead"
      Tab(4).Control(2)=   "Label17"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "&6. Copies"
      TabPicture(5)   =   "frmProduct1.frx":07AC
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grid1"
      Tab(5).Control(1)=   "cmdAddCopy"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "cmdRemove"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      Begin VB.Frame frPrices 
         Caption         =   "Prices"
         ForeColor       =   &H8000000D&
         Height          =   2580
         Left            =   645
         TabIndex        =   83
         Top             =   450
         Width           =   5190
         Begin VB.TextBox txtUSPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
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
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   91
            Top             =   675
            Width           =   1380
         End
         Begin VB.TextBox txtUKPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
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
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   90
            Top             =   285
            Width           =   1380
         End
         Begin VB.TextBox txtCost 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
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
            Left            =   3675
            Locked          =   -1  'True
            TabIndex        =   89
            Top             =   1440
            Width           =   1380
         End
         Begin VB.TextBox txtSP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
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
            Left            =   3675
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   675
            Width           =   1380
         End
         Begin VB.TextBox txtRRP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
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
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   87
            Top             =   285
            Width           =   1380
         End
         Begin VB.TextBox txtSSP 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
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
            Left            =   3675
            Locked          =   -1  'True
            TabIndex        =   86
            Top             =   1050
            Width           =   1380
         End
         Begin VB.CheckBox chkNDA 
            Caption         =   "No discount allowed"
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
            Height          =   315
            Left            =   150
            TabIndex        =   85
            Top             =   1635
            Width           =   2025
         End
         Begin VB.TextBox txtEUPrice 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
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
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   84
            Top             =   1080
            Width           =   1380
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "USD Price"
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
            TabIndex        =   98
            Top             =   720
            Width           =   870
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "GBP Price"
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
            TabIndex        =   97
            Top             =   330
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
            Left            =   2865
            TabIndex        =   96
            Top             =   1470
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
            Left            =   2865
            TabIndex        =   95
            Top             =   690
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
            Left            =   2865
            TabIndex        =   94
            Top             =   300
            Width           =   750
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Special S.P."
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
            Left            =   2565
            TabIndex        =   93
            Top             =   1080
            Width           =   1050
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Caption         =   "EUR Price"
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
            Left            =   90
            TabIndex        =   92
            Top             =   1125
            Width           =   870
         End
      End
      Begin VB.CommandButton cmdUnlockMB 
         BackColor       =   &H00C4BCA4&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6135
         Picture         =   "frmProduct1.frx":07C8
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   510
         Width           =   540
      End
      Begin VB.Frame frMultibuys 
         Caption         =   "Multibuys"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   855
         Left            =   6735
         TabIndex        =   79
         Top             =   390
         Visible         =   0   'False
         Width           =   2565
         Begin VB.ComboBox cboMultibuys 
            Appearance      =   0  'Flat
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
            Left            =   90
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   390
            Width           =   2340
         End
      End
      Begin VB.TextBox txtWeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   -71430
         TabIndex        =   77
         Top             =   540
         Width           =   1200
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H00C4BCA4&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74400
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   3210
         Width           =   405
      End
      Begin VB.CommandButton cmdAddCopy 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74850
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   3210
         Width           =   420
      End
      Begin VB.TextBox txtReason 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   1020
         Left            =   6060
         TabIndex        =   70
         Top             =   2175
         Visible         =   0   'False
         Width           =   4290
      End
      Begin VB.CommandButton cmdUnlockPrices 
         BackColor       =   &H00C4BCA4&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   165
         Picture         =   "frmProduct1.frx":0B52
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   555
         Width           =   435
      End
      Begin VB.CheckBox chkMAG 
         Caption         =   "Newspaper type"
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
         Left            =   -72465
         TabIndex        =   65
         Top             =   2745
         Width           =   1665
      End
      Begin VB.CheckBox chkObsolete 
         Caption         =   "Obsolete"
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
         Height          =   315
         Left            =   -72465
         TabIndex        =   64
         Top             =   2985
         Width           =   1245
      End
      Begin VB.Frame Frame4 
         Caption         =   "Catalogue entries"
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
         Height          =   3105
         Left            =   -66420
         TabIndex        =   59
         Top             =   390
         Width           =   2460
         Begin VB.CommandButton cmRemoveCat 
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
            Height          =   345
            Left            =   195
            Style           =   1  'Graphical
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   2655
            Width           =   945
         End
         Begin VB.ComboBox cboCATAL 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   285
            TabIndex        =   61
            Top             =   345
            Width           =   825
         End
         Begin VB.CommandButton cmdAddCat 
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
            Height          =   420
            Left            =   1110
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   330
            Width           =   945
         End
         Begin MSComctlLib.ListView lvwCE 
            Height          =   1890
            Left            =   105
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   750
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   3334
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cat. No."
               Object.Width           =   1658
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Price"
               Object.Width           =   2187
            EndProperty
         End
      End
      Begin VB.ComboBox cboCatHead 
         Appearance      =   0  'Flat
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
         Left            =   -74580
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   750
         Width           =   7500
      End
      Begin VB.Frame Frame1 
         Caption         =   "Publisher's status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2145
         Left            =   -68010
         TabIndex        =   53
         Top             =   690
         Visible         =   0   'False
         Width           =   3975
         Begin VB.PictureBox Picture 
            Height          =   1800
            Left            =   60
            ScaleHeight     =   1740
            ScaleWidth      =   2100
            TabIndex        =   100
            Top             =   315
            Width           =   2160
            Begin VB.OptionButton optRP 
               Caption         =   "Reprinting"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   210
               TabIndex        =   106
               Top             =   624
               Width           =   1335
            End
            Begin VB.OptionButton optOOP 
               Caption         =   "Out of print"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   210
               TabIndex        =   105
               Top             =   312
               Width           =   1335
            End
            Begin VB.OptionButton optIP 
               Caption         =   "In print"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   195
               TabIndex        =   104
               Top             =   0
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton optBO 
               Caption         =   "On backorder"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   225
               TabIndex        =   103
               Top             =   936
               Width           =   1620
            End
            Begin VB.OptionButton optMR 
               Caption         =   "Market restricted"
               ForeColor       =   &H8000000D&
               Height          =   270
               Left            =   210
               TabIndex        =   102
               Top             =   1485
               Width           =   1845
            End
            Begin VB.OptionButton optNYP 
               Caption         =   "Not yet printed"
               ForeColor       =   &H8000000D&
               Height          =   195
               Left            =   210
               TabIndex        =   101
               Top             =   1248
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdRediarize 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Re-diarize"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2250
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   630
            Width           =   1515
         End
         Begin VB.CommandButton cmdCancelPOL 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Cancel P.O.L.s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2250
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   1020
            Width           =   1515
         End
         Begin VB.CommandButton cmdView 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&View P.O.L.s"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2250
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.CommandButton cmdNewWant 
         BackColor       =   &H00C4BCA4&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66345
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   555
         Width           =   495
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1080
         Left            =   -74730
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   825
         Width           =   6570
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   915
         Left            =   -74730
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   2370
         Width           =   6600
      End
      Begin VB.ComboBox cboProductType 
         Appearance      =   0  'Flat
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
         Left            =   -66045
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   510
         Width           =   2115
      End
      Begin VB.Frame Frame3 
         Caption         =   "Categories"
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
         Height          =   2535
         Left            =   -67350
         TabIndex        =   38
         Top             =   960
         Width           =   3420
         Begin VB.CommandButton cmdRefresh 
            BackColor       =   &H00C4BCA4&
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2910
            Picture         =   "frmProduct1.frx":0EDC
            Style           =   1  'Graphical
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   285
            Width           =   375
         End
         Begin VB.CommandButton cmdMarkForWebExport 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Mark for Web export"
            Height          =   315
            Left            =   435
            Style           =   1  'Graphical
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   2145
            Width           =   2475
         End
         Begin VB.CommandButton cmdUP 
            BackColor       =   &H00C4BCA4&
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1830
            Width           =   330
         End
         Begin VB.ComboBox cboSection 
            Appearance      =   0  'Flat
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
            Left            =   420
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   285
            Width           =   2490
         End
         Begin VB.CommandButton cmdAddSection 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Add"
            Height          =   315
            Left            =   435
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   675
            Width           =   750
         End
         Begin VB.CommandButton cmdRemoveSection 
            BackColor       =   &H00C4BCA4&
            Caption         =   "&Remove"
            Height          =   315
            Left            =   2535
            Style           =   1  'Graphical
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   645
            Width           =   750
         End
         Begin MSComctlLib.ListView lvw 
            Height          =   1140
            Left            =   420
            TabIndex        =   42
            Top             =   1005
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   2011
            SortKey         =   1
            View            =   3
            SortOrder       =   -1  'True
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Section "
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Priority"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.TextBox txtLoyaltyRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   -73875
         TabIndex        =   36
         Top             =   1365
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Frame Frame2 
         Caption         =   "B.I.C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   -70005
         TabIndex        =   33
         Top             =   405
         Width           =   2535
         Begin VB.TextBox txtBICDescriptions 
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
            Height          =   510
            Left            =   135
            MultiLine       =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   690
            Width           =   2280
         End
         Begin VB.TextBox txtBIC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   120
            TabIndex        =   13
            Top             =   270
            Width           =   1380
         End
      End
      Begin VB.CommandButton cmdSupplier 
         BackColor       =   &H00C4BCA4&
         Caption         =   "· · ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -70860
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1920
         Width           =   480
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   -73860
         TabIndex        =   12
         Top             =   945
         Width           =   1185
      End
      Begin VB.CommandButton cmdSetDefault 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Default V.A.T. rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -72630
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   915
         Width           =   1755
      End
      Begin VB.TextBox txtBinding 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   -73860
         TabIndex        =   11
         Top             =   540
         Width           =   1200
      End
      Begin MSComctlLib.ListView lvwWants 
         Height          =   2565
         Left            =   -74790
         TabIndex        =   51
         Top             =   585
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   4524
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483635
         BackColor       =   14416635
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2152
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   3951
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Note"
            Object.Width           =   6068
         EndProperty
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   2685
         Left            =   -74865
         OleObjectBlob   =   "frmProduct1.frx":1266
         TabIndex        =   75
         Top             =   435
         Width           =   10965
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Weight"
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
         Left            =   -72570
         TabIndex        =   78
         Top             =   570
         Width           =   1080
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for price/discount alteration (min 10 chars.)"
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
         Left            =   6120
         TabIndex        =   69
         Top             =   1875
         Visible         =   0   'False
         Width           =   4350
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "See safe status"
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
         Left            =   -74880
         TabIndex        =   67
         Top             =   2925
         Width           =   810
      End
      Begin VB.Label lblSeesafe 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -74040
         TabIndex        =   66
         Top             =   2865
         Width           =   1095
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
         Height          =   285
         Left            =   -74460
         TabIndex        =   58
         Top             =   510
         Width           =   1635
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
         Left            =   -74730
         TabIndex        =   49
         Top             =   525
         Width           =   1035
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
         Left            =   -74700
         TabIndex        =   48
         Top             =   2100
         Width           =   885
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Product type"
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
         Left            =   -67125
         TabIndex        =   45
         Top             =   570
         Width           =   1035
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
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
         Left            =   -75015
         TabIndex        =   37
         Top             =   1395
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Deal"
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
         Left            =   -74910
         TabIndex        =   35
         Top             =   2415
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Supplier"
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
         Left            =   -74910
         TabIndex        =   32
         Top             =   1995
         Width           =   810
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
         Left            =   -75000
         TabIndex        =   28
         Top             =   975
         Width           =   1080
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
         Left            =   -75000
         TabIndex        =   27
         Top             =   570
         Width           =   1080
      End
      Begin VB.Label lblDeal 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   -74055
         TabIndex        =   25
         Top             =   2340
         Width           =   3135
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   360
         Left            =   -74025
         TabIndex        =   24
         Top             =   1935
         Width           =   3135
      End
   End
   Begin VB.TextBox txtPubPlace 
      Appearance      =   0  'Flat
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
      Left            =   6480
      TabIndex        =   5
      Top             =   510
      Width           =   2520
   End
   Begin VB.TextBox txtPubDate 
      Appearance      =   0  'Flat
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
      Left            =   6480
      TabIndex        =   6
      Top             =   885
      Width           =   2520
   End
   Begin VB.TextBox txtEdition 
      Appearance      =   0  'Flat
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
      Left            =   5520
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1635
      Width           =   3480
   End
   Begin VB.TextBox txtPublisher 
      Appearance      =   0  'Flat
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
      Left            =   6480
      MaxLength       =   255
      TabIndex        =   7
      Top             =   1260
      Width           =   2520
   End
   Begin VB.TextBox txtSubtitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   750
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1125
      Width           =   3900
   End
   Begin VB.TextBox txtAuthor 
      Appearance      =   0  'Flat
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
      Left            =   750
      MaxLength       =   255
      TabIndex        =   4
      Top             =   1755
      Width           =   3915
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   735
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   495
      Width           =   3900
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
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
      Left            =   1470
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label lblHelp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   6480
      TabIndex        =   52
      Top             =   30
      Width           =   360
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN-13 / E.A.N."
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
      Left            =   3210
      TabIndex        =   29
      Top             =   150
      Width           =   1470
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
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
      Left            =   9090
      TabIndex        =   23
      Top             =   510
      Width           =   420
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   4845
      TabIndex        =   21
      Top             =   525
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   4950
      TabIndex        =   20
      Top             =   915
      Width           =   1470
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   4815
      TabIndex        =   19
      Top             =   1695
      Width           =   645
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   5445
      TabIndex        =   18
      Top             =   1305
      Width           =   975
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
      Left            =   15
      TabIndex        =   17
      Top             =   1800
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
      Left            =   15
      TabIndex        =   16
      Top             =   1140
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
      Left            =   165
      TabIndex        =   15
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN-10 / #code"
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
      TabIndex        =   14
      Top             =   165
      Width           =   1350
   End
End
Attribute VB_Name = "frmProduct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim flgLoading As Boolean
Dim mCancel As Boolean
Dim XA As XArrayDB
Dim frmPrevious As Form
Dim tlCATAL As z_TextList
Dim bPriceChange As Boolean
Dim struct_OldPrices As OldPrices
Dim strPriceChangeReason As String
Dim lngSMIDPriceChange As Long
Dim PID As String

Private Sub chkCore_Click()
    On Error GoTo errHandler

    If flgLoading Then Exit Sub
    If chkCore = 1 Then
        oProd.SetCore True
    Else
        oProd.SetCore False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.chkCore_Click", , EA_NORERAISE
'    HandleError
'
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.chkCore_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkNDA_Click()
    On Error GoTo errHandler

    If flgLoading Then Exit Sub
    If chkNDA = 1 Then
        oProd.SetNDA True
    Else
        oProd.SetNDA False
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.chkNDA_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.chkNDA_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdAddCat_Click()
    On Error GoTo errHandler
Dim oCE As a_CATALP
    If cboCATAL = "" Then Exit Sub
    Set oCE = oProd.CatalogueEntries.Add
    oCE.BeginEdit
    oCE.CATALID = tlCATAL.Key(Me.cboCATAL)
    oCE.Serial = cboCATAL
    oCE.Price = oProd.SP
    oCE.ApplyEdit
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdAddCat_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdChangeType_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim strMsg As String
Dim frmPreview As frmProductPrev
    If MsgBox("You want to change this product to be a general product (non book)?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    oProd.SetProductType "G"
    oProd.ApplyEdit lngResult, strMsg
    If lngResult = 99 Then
        WaitMsg "", False, Me
        If strMsg = "TIMEOUT" Then
            MsgBox "The operation has timed out. The record is probably locked by another user." & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
        ElseIf strMsg = "DUPLICATE" Then
            MsgBox "Invalid values - check that the code is has not been already used", vbInformation, "Save failed"
        End If
    Else
        If frmPreview Is Nothing Then
            Set frmPreview = New frmProductPrev
        Else
            Set frmPreview = frmPrevious
        End If
        frmPreview.component oProd
        frmPreview.RefreshForm
        frmPreview.Show
        WaitMsg "", False, Me
        Unload Me
    End If
    
    Unload Me
    
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdChangeType_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdChangeType_Click", , EA_NORERAISE
    HandleError
End Sub

Sub component(pProduct As a_Product, Optional pPrevForm As Form)
    On Error GoTo errHandler
    Set frmPrevious = pPrevForm
    Set oProd = pProduct
    oProd.BeginEdit
    oProd.SetBook
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.Component(pProduct,pPrevForm)", Array(pProduct, pPrevForm)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.component(pProduct,pPrevForm)", Array(pProduct, pPrevForm)
End Sub

Sub cmdHelp_Click()
    On Error GoTo errHandler
Dim frm As New frmHelpGen
Dim tmp As String
    tmp = LoadResString(101)
    frm.component tmp, "ISBN-10,ISBN-13 and EAN codes", 8000, 3000
    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdHelp_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboProductType_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetProductTypeID oPC.Configuration.ProductTypes.Key(cboProductType)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cboProductType_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboMultibuys_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetMultibuyCode oPC.Configuration.Multibuys.f4(oPC.Configuration.Multibuys.Key(cboMultibuys))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cboMultibuys_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cboCatHead_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetCatalogueheadingID oPC.Configuration.CatalogueHeadings.Key(cboCatHead)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cboCatHead_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cboCatHead_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub chkServiceItem_Click()
'    If chkServiceItem Then
'        oProd.SetServiceItem
'    Else
'        oProd.SetBook
'    End If
'End Sub

Private Sub chkObsolete_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.Obsolete = chkObsolete
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.chkObsolete_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.chkObsolete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkMAG_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If chkMAG = 1 Then
        oProd.SetMagsEtc
    Else
        oProd.SetBook
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.chkMAG_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.chkMAG_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddCopy_Click()
    On Error GoTo errHandler
Dim frm As frmCopy
Dim oCopy As a_Copy
Dim tmpCopy As a_Copy
    Set frm = New frmCopy
    Set oCopy = oProd.Copies.Add
    
    If Grid1.Bookmark > 0 Then
        Set tmpCopy = oProd.Copies(Grid1.Bookmark)
    Else
        Set tmpCopy = Nothing
    End If
    
    
    frm.component oCopy, tmpCopy
    frm.Show vbModal
    Set oCopy = Nothing
    Set frm = Nothing
    LoadCopies
    Grid1.ReBind
    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdAddCopy_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdAddCopy_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdDelete_Click()
    On Error GoTo errHandler

    If MsgBox("Confirm deletion of product: " & oProd.Title & vbCrLf & "This is permanent!", vbOKCancel + vbExclamation, "Confirm") = vbOK Then
        oProd.Delete
        oProd.ApplyEdit
        Unload Me
    End If


'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdDelete_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddSection_Click()
    On Error GoTo errHandler
Dim oPSEC As New a_ProductSection
    If flgLoading Then Exit Sub
    If cboSection = "" Then
        MsgBox "You cannot add an empty section description.", vbInformation, "Can't do this"
        Exit Sub
    End If
    If InStr(1, cboSection, "Unallocated") > 0 Then
        MsgBox "You cannot add to the 'Unallocated' section.", vbInformation, "Can't do this"
        Exit Sub
    End If
    
    Set oPSEC = oProd.ProductSections.Add
    oPSEC.PID = oProd.PID
    oPSEC.SECID = oPC.Configuration.Sections.Key(cboSection)
    oPSEC.Description = cboSection
    If oProd.ProductSections.Count = 0 Or oProd.ProductSections.Count = 1 And oProd.MultibuyCode > "" Then
        oPSEC.Priority = 99
        oProd.MasterCategory = oPSEC.SECID
    End If
    oPSEC.ApplyEdit
    oPSEC.BeginEdit
    cboSection.RemoveItem cboSection.ListIndex
    If cboSection.ListCount > 0 Then
        cboSection.ListIndex = 0
    Else
        cboSection.ListIndex = -1
    End If
    LoadPSECs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdAddSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMarkForWebExport_Click()
    On Error GoTo errHandler
Dim oPSEC As New a_ProductSection
    If flgLoading Then Exit Sub
    Set oPSEC = oProd.ProductSections.Add
    oPSEC.PID = oProd.PID
    oPSEC.SECID = oPC.Configuration.WebExportID
    oPSEC.Description = "For Web export"
    oPSEC.ApplyEdit
    oPSEC.BeginEdit
    If cboSection.ListCount > 0 Then
        cboSection.ListIndex = 0
    Else
        cboSection.ListIndex = -1
    End If
    LoadPSECs
    Exit Sub

'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdMarkForWebExport_Click"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdMarkForWebExport_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefresh_Click()
    
    oPC.Configuration.ReloadCategories
    LoadCombo cboSection, oPC.Configuration.Sections_Short
    If cboSection.ListCount > 0 Then
        cboSection.ListIndex = 0
    Else
        cboSection.ListIndex = -1
    End If

End Sub

Private Sub cmdRemoveSection_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If lvw.ListItems.Count = 0 Then Exit Sub
    If Not oProd.ProductSections.Remove(oProd.ProductSections.Key(lvw.SelectedItem)) Then
        MsgBox "Cannot remove this category assignment, possibly it is the master category. First assign a new master category.", vbInformation + vbOKOnly, "Can't do this"
        Exit Sub
    End If
    If oPC.Configuration.Sections.Key(lvw.SelectedItem) <> 0 Then   'only if not a 'system' category like 'for web export'
        cboSection.AddItem lvw.SelectedItem
        cboSection.ListIndex = 0
        LoadPSECs
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdRemoveSection_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdUP_Click()
    On Error GoTo errHandler
    If lvw.SelectedItem Is Nothing Then Exit Sub
    If oProd.ProductSections.Item(lvw.SelectedItem.Key).SECID <> oPC.Configuration.WebExportID And _
            InStr(1, lvw.SelectedItem, "Multibuy") = 0 Then
        oProd.ProductSections.Mark oProd.ProductSections.Key(lvw.SelectedItem)
        LoadPSECs
    Else
        MsgBox "You cannot assign a priority category to this category.", vbInformation, "Can't do this"
    End If
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProduct: cmdUP_Click"  'unknown source
        If errRepeat < 5 Then
            Err.Clear
            Exit Sub
        Else
            LogSaveToFile "Access violation in frmProduct: cmdUP_Click after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdUP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadPSECs()
    On Error GoTo errHandler
Dim lstItem As ListItem
Dim i As Long
    
    lvw.ListItems.Clear
    For i = 1 To oProd.ProductSections.Count
        Set lstItem = lvw.ListItems.Add
        With oProd.ProductSections(i)
            lstItem.text = .Description
            If lstItem.Key = "" Then lstItem.Key = .Key
            lstItem.SubItems(1) = .PriorityF
        End With
    Next i
    
EXIT_Handler:
    Set lstItem = Nothing
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.LoadPSECs"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.LoadPSECs"
End Sub

'Private Sub LoadMultibuys()
'    On Error GoTo errHandler
'Dim lstItem As ListItem
'Dim i As Long
'
'    lvw.ListItems.Clear
'    For i = 1 To oProd.ProductSections.Count
'        Set lstItem = lvw.ListItems.Add
'        With oProd.ProductSections(i)
'            lstItem.Text = .Description
'            If lstItem.key = "" Then lstItem.key = .key
'        End With
'    Next i
'
'EXIT_Handler:
'    Set lstItem = Nothing
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.LoadMultibuys"
'End Sub

'Private Sub cmdGenerateEAN_Click()
'    On Error GoTo errHandler
'Dim oProdCode As New z_ProdCode
'    oProdCode.SetCodesForBook txtCode
'    oProd.SetEAN oProdCode.Ean
'    txtEAN = oProd.Ean
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdGenerateEAN_Click", , EA_NORERAISE
'    HandleError
'End Sub

Private Sub cmdNewWant_Click()
    On Error GoTo errHandler
Dim frm As frmWant
'Dim oWant As a_Want
   ' Set oWant = oProd.Wants.Add
    Set frm = New frmWant
    frm.component oProd
    frm.Show vbModal
    LoadWants
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdNewWant_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdNewWant_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRediarize_Click()
    On Error GoTo errHandler
Dim frm As New frmRediarize
Dim oSM As New z_StockManager

    frm.Show vbModal
    If Not frm.Cancelled Then
        oSM.RediarizePOLS oProd.PID, frm.RediarizedPeriod, frm.Reason
        If frm.Reason = "R" Then
            oProd.SetStatus enAwaitingReprint
        End If
    End If
    Unload frm
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdRediarize_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdRediarize_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancelPOL_Click()
    On Error GoTo errHandler
Dim frm As New frmCancelPOLs
Dim oSM As New z_StockManager

    frm.Show vbModal
    If Not frm.Cancelled Then
        oSM.CANCELPOLs oProd.PID, frm.Reason
        Select Case frm.Reason
        Case "O"
            oProd.SetStatus enOutOfPrint
        Case "R"
            oProd.SetStatus enAwaitingReprint
        End Select
    End If
    Unload frm
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdCancelPOL_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdCancelPOL_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo errHandler
Dim oCopy As a_Copy
    Set oCopy = oProd.Copies(XA(Grid1.Bookmark, 6))
    oCopy.BeginEdit
    oCopy.Delete
    oCopy.ApplyEdit
    LoadCopies
    Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdRemove_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdRemove_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub UPdateFromXArray()
    On Error GoTo errHandler
'Dim i As Long
'Dim strMsg As String
'    strMsg = ""
'    For i = 1 To XA.UpperBound(1)
'        If Not oProd.Copies(XA(i, 6)).IsValid Then
'            strMsg = strMsg & IIf(Len(strMsg) > 0, vbCrLf, "") & "Row"
'        End If
'    Next
'    For i = 1 To XA.UpperBound(1)
'        If oProd.Copies(XA(i, 6)).IsValid Then
'            oProd.Copies(XA(i, 6)).BeginEdit
'            oProd.Copies(XA(i, 6)).SetComment XA.Value(i, 2)
'            oProd.Copies(XA(i, 6)).SetPurchaseDate XA.Value(i, 3)
'            oProd.Copies(XA(i, 6)).SetSoldDate XA.Value(i, 4)
'            oProd.Copies(XA(i, 6)).SetPrice XA.Value(i, 5)
'            oProd.Copies(XA(i, 6)).ApplyEdit
'        End If
'    Next
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.UPdateFromXArray"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.UPdateFromXArray"
End Sub

Private Sub cmdRemoveWant_Click()
    On Error GoTo errHandler
'Dim oWant As a_Want
'    Set oWant = oProd.Wants(lvwWants.SelectedItem.Key)
'    oWant.BeginEdit
'    oWant.Delete
'    oWant.ApplyEdit
'    LoadWants
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdRemoveWant_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdRemoveWant_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSetDefault_Click()
    On Error GoTo errHandler
    oProd.VATRate = oPC.Configuration.VATRate
    Me.txtVAT = PBKSPercentF(oPC.Configuration.VATRate)
    mSetfocus txtVAT
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdSetDefault_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdSetDefault_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdSupplier_Click()
    On Error GoTo errHandler
Dim frm As New frmBrowseSUppliers2
    frm.Show vbModal
    If frm.SupplierID > 0 Then
        oProd.SupplierID = frm.SupplierID
        oProd.LastSupplierName = frm.SupplierName
        Me.lblSupplier = oProd.LastSupplierName
    Else
        MsgBox "No supplier selected.", vbOKOnly, "Warning"
    End If
    Unload frm
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdSupplier_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdSupplier_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdUnlockMB_Click()
    On Error GoTo errHandler
Dim bIsSUpervisor As Boolean
Dim bCancelled As Boolean
Dim strName As String
    If SecurityControl(enSECURITY_MULTIBUY, bCancelled, "Enter your signature", "You do not have permission to alter multibuy allocations (or your signature is invalid)", bIsSUpervisor, strName, lngSMIDPriceChange) = True Then
        bPriceChange = True
        struct_OldPrices.Cost = oProd.Cost
        struct_OldPrices.SP = oProd.SP
        struct_OldPrices.RRP = oProd.RRP
        struct_OldPrices.SpecialPrice = oProd.SpecialPrice
        struct_OldPrices.UKPrice = oProd.UKPrice
        struct_OldPrices.USPrice = oProd.USPrice
        struct_OldPrices.EUPrice = oProd.EUPrice
        struct_OldPrices.MultibuyCode = oProd.MultibuyCode
        struct_OldPrices.NDA = oProd.IsNDA
        Me.cboMultibuys.Locked = False
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdUnlockMB_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdUnlockPrices_Click()
    On Error GoTo errHandler
Dim bIsSUpervisor As Boolean
Dim bCancelled As Boolean
Dim strName As String
Dim frm As New frmAudit_ProductPrices

    If SecurityControl(enSECURITY_CUSTDISCOUNT_AUTH, bCancelled, "Enter your signature", "You do not have permission to unlock the price fields (or your signature is invalid)", bIsSUpervisor, strName, lngSMIDPriceChange) = True Then
        bPriceChange = True
        'Set old prices
        struct_OldPrices.Cost = oProd.Cost
        struct_OldPrices.SP = oProd.SP
        struct_OldPrices.RRP = oProd.RRP
        struct_OldPrices.SpecialPrice = oProd.SpecialPrice
        struct_OldPrices.UKPrice = oProd.UKPrice
        struct_OldPrices.USPrice = oProd.USPrice
        struct_OldPrices.EUPrice = oProd.EUPrice
        struct_OldPrices.MultibuyCode = oProd.MultibuyCode
        struct_OldPrices.NDA = oProd.IsNDA
        
        frm.Show vbModal
        strPriceChangeReason = frm.Reason
        If Not frm.Cancelled Then
            frPrices.Enabled = True
            Me.txtRRP.Locked = False
            Me.txtRRP.BackColor = &H80000005
            Me.txtSP.Locked = False
            Me.txtSP.BackColor = &H80000005
            Me.txtSSP.Locked = False
            Me.txtSSP.BackColor = &H80000005
            Me.txtCost.Locked = False
            Me.txtCost.BackColor = &H80000005
            Me.txtUKPrice.Locked = False
            Me.txtUKPrice.BackColor = &H80000005
            Me.txtEUPrice.Locked = False
            Me.txtEUPrice.BackColor = &H80000005
            Me.txtUSPrice.Locked = False
            Me.txtUSPrice.BackColor = &H80000005
            Me.chkNDA.Enabled = True
            Me.frMultibuys.Enabled = True
            Me.cboMultibuys.Locked = False
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdUnlockPrices_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdView_Click()
    On Error GoTo errHandler
Dim frm As New frmPOLsPerPID_OS
    frm.component oProd.PID
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdView_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemoveCAt_Click()
    On Error GoTo errHandler
Dim oCE As a_CATALP
    Set oCE = oProd.CatalogueEntries(lvwCE.SelectedItem.Key)
    oCE.BeginEdit
    oCE.Delete
    oCE.ApplyEdit
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdRemoveCAt_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub cmRemoveCat_Click()
    On Error GoTo errHandler
Dim oCE As a_CATALP
    Set oCE = oProd.CatalogueEntries(lvwCE.SelectedItem.Key)
    oCE.BeginEdit
    oCE.Delete
    oCE.ApplyEdit
    LoadListView
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmRemoveCat_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oProd.IsEditing Then oProd.CancelEdit
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.Form_Unload(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim i As Integer
Dim oC As a_Copy
    i = ColIndex + 1
    Set oC = oProd.Copies(XA(Grid1.Bookmark, 7))
    oC.BeginEdit
    Select Case i
    Case 2
        If Not oC.SetDescription(Grid1.text) Then
            OldValue = Grid1.text
            Cancel = True
        End If
    Case 3
        If Not oC.SetComment(Grid1.text) Then
            OldValue = Grid1.text
            Cancel = True
        End If
    Case 4
        If Not oC.SetPurchaseDate(Grid1.text) Then
            OldValue = Grid1.text
            Cancel = True
        End If
    Case 5
        If Not oC.SetSoldDate(Grid1.text) Then
            OldValue = Grid1.text
            Cancel = True
        End If
    Case 6
        If Not oC.SetPrice(Grid1.text) Then
            OldValue = Grid1.text
            Cancel = True
        End If
    End Select
    If Err Then
        OldValue = Grid1.text
        Cancel = True
    End If

    If Cancel = True Then
        oC.CancelEdit
    Else
        oC.ApplyEdit
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
'         Cancel), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
'MsgBox "Selected row is : " & Grid1.Row + 1
Dim frm As frmCopy
Dim oCopy As a_Copy
Dim tmpCopy As a_Copy
    If IsNull(Grid1.Bookmark) Then
        Exit Sub
    End If
    Set oCopy = oProd.Copies(XA(Grid1.Bookmark, 7))
    Set frm = New frmCopy
    If Grid1.Bookmark > 0 Then
        Set tmpCopy = oProd.Copies(Grid1.Bookmark)
    Else
        Set tmpCopy = Nothing
    End If

    frm.component oCopy, tmpCopy
    frm.Show vbModal

    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProduct: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmProduct: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 5) > "" Then
        RowStyle.BackColor = &HDCDBF2
    End If
    If XA(Bookmark, 8) = True Then
        RowStyle.BackColor = vbRed
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
'         RowStyle), EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub lblHelp_DblClick()
    On Error GoTo errHandler
Dim str As String
    str = "If you have an ISBN-13 or an EAN code, enter it in the field labelled 'ISBN-13/EAN'" & vbCrLf
    str = str & "If you have an ISBN-10 code only, then enter it in the field labelled 'ISBN-10/#code'" & vbCrLf
    str = str & "If you have neither, then either enter a code of your choice (e.g. #CARD10) in the 'ISBN-10/#code' field" & vbCrLf
    str = str & "   or place a '#' symbol in the field and Papyrus will generate a unique code and EAN for the product."
    MsgBox str, vbInformation + vbOKOnly, "Hints"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.lblHelp_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Lvw_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.lvw_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwWants_DblClick()
    On Error GoTo errHandler
'Dim frm As frmWant
'Dim oWant As a_Want
'    Set oWant = oProd.Wants(lvwWants.SelectedItem.Key)
'    Set frm = New frmWant
'    frm.Component oProd.pID
'    frm.Show vbModal
'    LoadWants
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.lvwWants_DblClick", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.lvwWants_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub oProd_RedisplayCodes()
    On Error GoTo errHandler
    txtCode = oProd.code
    txtEAN = oProd.EAN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.oProd_RedisplayCodes", , EA_NORERAISE
    HandleError
End Sub

Private Sub oProd_Valid(strMsg As String)
    On Error GoTo errHandler
    Me.txtErrors = strMsg
    Me.cmdOK.Enabled = (strMsg = "")
    Me.cmdAddCopy.Enabled = (strMsg = "")
    Me.cmdRemove.Enabled = (strMsg = "")
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.oProd_Valid(strMsg)", strMsg, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.oProd_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    If oProd.IsEditing Then oProd.CancelEdit
    Unload Me
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdCancel_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdNewCode_Click()
    On Error GoTo errHandler
    Me.txtCode = "#"
    oProd.SetCode "#"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdNewCode_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdNewCode_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim strMsg As String
Dim frmPreview As frmProductPrev
Dim oPT As a_PT
Dim oSM As z_StockManager

    If oPC.GetProperty("SecureTF") = "TRUE" Then
        If bPriceChange = True Then  'WRite and audit record
            
            Set oSM = New z_StockManager
            If struct_OldPrices.SP <> oProd.SP Then
                oSM.InsertAuditRecord "SP", strPriceChangeReason, Format(struct_OldPrices.SP / 100, "###,##0.00"), Format(oProd.SP / 100, "###,##0.00"), lngSMIDPriceChange, oProd.PID, 0
            End If
            If struct_OldPrices.RRP <> oProd.RRP Then
                oSM.InsertAuditRecord "RRP", strPriceChangeReason, Format(struct_OldPrices.RRP / 100, "###,##0.00"), Format(oProd.RRP / 100, "###,##0.00"), lngSMIDPriceChange, oProd.PID, 0
            End If
            If struct_OldPrices.Cost <> oProd.Cost Then
                oSM.InsertAuditRecord "COST", strPriceChangeReason, Format(struct_OldPrices.Cost / 100, "###,##0.00"), Format(oProd.Cost / 100, "###,##0.00"), lngSMIDPriceChange, oProd.PID, 0
            End If
            If struct_OldPrices.SpecialPrice <> oProd.SpecialPrice Then
                oSM.InsertAuditRecord "SSP", strPriceChangeReason, Format(struct_OldPrices.SpecialPrice / 100, "###,##0.00"), Format(oProd.SpecialPrice / 100, "###,##0.00"), lngSMIDPriceChange, oProd.PID, 0
            End If
            If struct_OldPrices.MultibuyCode <> oProd.MultibuyCode Then
                oSM.InsertAuditRecord "MB", strPriceChangeReason, struct_OldPrices.MultibuyCode, oProd.MultibuyCode, lngSMIDPriceChange, oProd.PID, 0
            End If
            If struct_OldPrices.NDA <> oProd.IsNDA Then
                oSM.InsertAuditRecord "NDA", strPriceChangeReason, CStr(struct_OldPrices.NDA), CStr(oProd.IsNDA), lngSMIDPriceChange, oProd.PID, 0
            End If
        End If
    End If
    WaitMsg "Saving product . . .", True, Me
    If oPC.AllowGeneralStock = False Then oProd.SetBook  'If we allow gneral stock then the other form is used for non-books - this is always books
    If oProd.SP = 0 Then
        Set oPT = New a_PT
        oPT.Load oProd.ProductTypeID
        oProd.SetSPFROMRRP oPT
        Set oPT = Nothing
    End If
    oProd.ApplyEdit lngResult, strMsg
    If lngResult = 99 Then
        WaitMsg "", False, Me
        If strMsg = "TIMEOUT" Then
            MsgBox "The operation has timed out. The record is probably locked by another user." & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
        ElseIf strMsg = "DUPLICATE" Then
            MsgBox "Invalid values - check that the code is has not been already used", vbInformation, "Save failed"
        ElseIf Left(strMsg, 5) = "ERROR" Then
            MsgBox "Invalid code values: " & strMsg, vbInformation, "Save failed"
        End If
    Else
        If frmPrevious Is Nothing Then
            Set frmPreview = New frmProductPrev
        Else
            Set frmPreview = frmPrevious
        End If
        frmPreview.component oProd
        frmPreview.RefreshForm
        frmPreview.Show
        WaitMsg "", False, Me
        Unload Me
    End If
'errHandler:
'    ErrPreserve
'    If err = 10005 Then Resume Next  'assume that this is the elusive vbcsExceptionFilter error that seems both harmless and untraceable
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdOK_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11500
        Height = 7000
    End If
    LoadControls
    Me.SSTab1.Tab = 0
    SSTab1.TabVisible(5) = oPC.Configuration.AllowCopyInfo
    SSTab1.TabVisible(4) = oPC.SupportsCatalogue
    SSTab1.TabVisible(3) = oPC.Configuration.AllowCopyInfo
    If oPC.GetProperty("SecureTF") = "TRUE" Then
        cmdUnlockPrices.Visible = True
        cmdUnlockMB.Visible = oPC.SupportsMultibuys
        Me.txtRRP.Locked = True
        Me.txtRRP.BackColor = &H80000018
        Me.txtSP.Locked = True
        Me.txtSP.BackColor = &H80000018
        Me.txtSSP.Locked = True
        Me.txtSSP.BackColor = &H80000018
        Me.txtCost.Locked = True
        Me.txtCost.BackColor = &H80000018
        Me.txtUKPrice.Locked = True
        Me.txtUKPrice.BackColor = &H80000018
        Me.txtEUPrice.Locked = True
        Me.txtEUPrice.BackColor = &H80000018
        Me.txtUSPrice.Locked = True
        Me.txtUSPrice.BackColor = &H80000018
        Me.chkNDA.Enabled = False
        Me.cboMultibuys.Locked = True
    Else
        cmdUnlockPrices.Visible = False
        cmdUnlockMB.Visible = False
        Me.txtRRP.Locked = False
        Me.txtRRP.BackColor = &H80000005
        Me.txtSP.Locked = False
        Me.txtSP.BackColor = &H80000005
        Me.txtSSP.Locked = False
        Me.txtSSP.BackColor = &H80000005
        Me.txtCost.Locked = False
        Me.txtCost.BackColor = &H80000005
        Me.txtUKPrice.Locked = False
        Me.txtUKPrice.BackColor = &H80000005
        Me.txtEUPrice.Locked = False
        Me.txtEUPrice.BackColor = &H80000005
        Me.txtUSPrice.Locked = False
        Me.txtUSPrice.BackColor = &H80000005
        Me.chkNDA.Enabled = True
        Me.cboMultibuys.Locked = False
    End If
    Set tlCATAL = New z_TextList
    tlCATAL.Load ltCatalogue
    LoadCombo Me.cboCATAL, tlCATAL
    LoadListView
    Me.chkMAG.Visible = (Not oPC.AllowGeneralStock)
    AutoSizeDropDownWidth Me.cboSection
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub RestrictCustomerTypes()
    On Error GoTo errHandler
Dim oPSEC As a_ProductSection
Dim i As Integer

    For Each oPSEC In oProd.ProductSections
        For i = cboSection.ListCount To 1 Step -1
            cboSection.ListIndex = i - 1
            If oPSEC.Description = cboSection Then
                cboSection.RemoveItem cboSection.ListIndex
            End If
        Next
    Next
    If cboSection.ListCount > 0 Then
        cboSection.ListIndex = 0
    Else
        cboSection.ListIndex = -1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.RestrictCustomerTypes"
End Sub

Private Sub LoadControls()
    On Error GoTo errHandler
Dim strPos As String
    flgLoading = True
    frMultibuys.Visible = oPC.SupportsMultibuys
    cmdUnlockMB.Visible = oPC.SupportsMultibuys
    txtCode = oProd.code
    txtEAN = oProd.EAN
    txtTitle = oProd.Title
    txtSubtitle = oProd.SubTitle
    txtAuthor = oProd.Author
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    txtPubPlace = oProd.PublicationPlace
    txtPubDate = oProd.PublicationDate
    txtBinding = oProd.BindingCode
    txtDescription = oProd.Description
    txtComment = oProd.Comment
    txtRRP = oProd.RRPF
    txtSP = oProd.SPF
    txtSSP = oProd.SpecialPriceF
    txtUKPrice = oProd.UKPriceF
    txtUSPrice = oProd.USPriceF
    txtEUPrice = oProd.EUPriceF
    txtCost = oProd.CostF
    txtWeight = oProd.WeightF
    txtBIC = oProd.BIC
    txtBICDescriptions = oPC.Configuration.BICs.FetchBICDescriptionsFromCodeSet(txtBIC)
    txtNote = oProd.Note
    txtVAT = oProd.VATRateF
    txtLoyaltyRate = oProd.LoyaltyRateF
strPos = "1"
    LoadCombo cboCatHead, oPC.Configuration.CatalogueHeadings
    LoadCombo cboSection, oPC.Configuration.Sections_Short
    LoadCombo cboProductType, oPC.Configuration.ProductTypes_Short
    'If oPC.SupportsMultibuys Then
        LoadCombo Me.cboMultibuys, oPC.Configuration.Multibuys
strPos = "2"
    On Error Resume Next
    cboProductType = oPC.Configuration.ProductTypes.Item(oProd.ProductTypeID)
    On Error GoTo errHandler
strPos = "3"
    If oProd.CatalogueheadingID > 0 Then cboCatHead = oPC.Configuration.CatalogueHeadings.Item(oProd.CatalogueheadingID)
strPos = "4"
    chkMAG = IIf(oProd.IsMagsEtc = True, 1, 0)
strPos = "5"
    chkObsolete = IIf(oProd.Obsolete = True, 1, 0)
    chkNDA = IIf(oProd.IsNDA = True, 1, 0)
    chkCore = IIf(oProd.IsCore, 1, 0)

strPos = "6"
    lblSupplier.Caption = oProd.LastSupplierName
strPos = "7"
    lblDeal.Caption = oProd.LastDealDescription
strPos = "8"
    
    Select Case oProd.Status
    Case "O"
        optOOP.Value = True
    Case "R"
        optRP.Value = True
    Case "B"
        optBO.Value = True
    Case "N"
        optNYP.Value = True
    Case "M"
        optMR.Value = True
    Case Else
        optIP.Value = True
    End Select
strPos = "9"
    LoadCopies
    LoadWants
    LoadPSECs
    SetMultibuyCode
    RestrictCustomerTypes
    Me.cmdChangeType.Visible = oPC.AllowGeneralStock
    
    flgLoading = False
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.LoadControls", , , , strPos, Array(strPos)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.LoadControls", , , , "strPos", Array(strPos)
End Sub
Private Sub SetMultibuyCode()
    On Error GoTo errHandler
    If oProd.MultibuyCode > "" Then
        Me.cboMultibuys = oPC.Configuration.Multibuys.ItemByF4(oProd.MultibuyCode)
    Else
        cboMultibuys = "<N/A>"
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.SetMultibuyCode"
End Sub
Private Sub LoadCopies()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String
Dim strCatalogues As String

    
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
 '   Grid1.ReBind
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.LoadCopies"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.LoadCopies"
End Sub
Private Sub LoadWants()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String
Dim strCatalogues As String
    lvwWants.ListItems.Clear
    For i = 1 To oProd.Wants.Count
        Set objItm = Me.lvwWants.ListItems.Add
        With objItm
            .Key = oProd.Wants(i).COLID & "k"
            .text = oProd.Wants(i).WantDateF
            .SubItems(1) = oProd.Wants(i).CustName
            .SubItems(2) = oProd.Wants(i).Note
        End With
    Next i
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.LoadWants"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.LoadWants"
End Sub

Private Sub lvwCopies_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.lvwCopies_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.lvwCopies_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub optIP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enInPrint
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.optIP_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.optIP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optMR_Click()
    On Error GoTo errHandler
    oProd.SetStatus enMarketRestricted
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.optMR_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optNYP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enNotYetPrinted

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.optNYP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOOP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enOutOfPrint
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.optOOP_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.optOOP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub optBO_Click()
    On Error GoTo errHandler
    oProd.SetStatus enOnBackorder
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.optBO_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.optBO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optRP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enAwaitingReprint
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.optRP_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.optRP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtBIC_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtBIC = oProd.BIC
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtBIC_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtBIC_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBIC_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    txtBICDescriptions = oPC.Configuration.BICs.FetchBICDescriptionsFromCodeSet(txtBIC)
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtBIC_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtBIC_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBIC_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetBIC(txtBIC)
    If Err Then
      Beep
      intPos = txtBIC.SelStart
      txtBIC = oProd.BIC
      txtBIC.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtBIC_Change", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtTitle_Change()
    On Error Resume Next
Dim intPos As Integer
    mCancel = Not oProd.SetTitle(txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oProd.Title
      txtTitle.SelStart = intPos - 1
    End If

End Sub

Private Sub txtWeight_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtWeight = oProd.Weight
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtWeight_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtWeight_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtWeight_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtWeight_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtWeight_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetWeight(txtWeight)
    If Err Then
      Beep
      intPos = txtWeight.SelStart
      txtWeight = oProd.Weight
      txtWeight.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtWeight_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtWeight_Change", , EA_NORERAISE
    HandleError
End Sub

'''



Private Sub txtBinding_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtBinding = oProd.BindingCode
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtBinding_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtBinding_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBinding_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtBinding_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtBinding_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBinding_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetBindingCode(txtBinding)
    If Err Then
      Beep
      intPos = txtBinding.SelStart
      txtBinding = oProd.BindingCode
      txtBinding.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtBinding_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtBinding_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim Res As Boolean
On Error Resume Next
Dim intPos As Integer

    oProd.SetCode (txtCode)
    If Len(oProd.code) > 10 Then
        MsgBox "The code you entered is greater than ten characters long. This is possibly an error." & vbCrLf _
                & "Make sure you are not entering an ISBN-13 or EAN number in this field. They belong in the next box.", vbExclamation + vbOKOnly, "Warning"
    End If
    If Err Then
        Beep
        intPos = txtCode.SelStart
        txtCode = oProd.code
        txtCode.SelStart = intPos - 1
    End If
    If txtCode <> "" Then
        If oProd.IsNew Then
            Res = oProd.Exists(txtCode)
            If Res = True Then
                MsgBox "This code is already being used. (" & txtCode & ")", vbInformation + vbOKOnly, "Invalid Code"
            End If
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtEAN_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtEAN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEAN_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEAN_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim Res As Boolean
Dim intPos As Integer

On Error Resume Next
    oProd.SetEAN (txtEAN)
    If Err Then
        Beep
        intPos = txtEAN.SelStart
        txtEAN = oProd.EAN
        txtEAN.SelStart = intPos - 1
    End If
    If txtEAN <> "" Then
        Res = oProd.Exists(txtEAN)
        If Res = True And oProd.IsNew Then
            MsgBox "This code is already being used. (" & txtEAN & ")", vbInformation + vbOKOnly, "Invalid EAN"
        End If
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEAN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub txtLoyaltyRate_GotFocus()
    On Error GoTo errHandler
    txtLoyaltyRate = oProd.LoyaltyRate
    AutoSelect txtLoyaltyRate
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtLoyaltyRate_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtLoyaltyRate_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtLoyaltyRate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.LoyaltyRate = txtLoyaltyRate
    
    txtLoyaltyRate = oProd.LoyaltyRateF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtLoyaltyRate_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtLoyaltyRate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtReason_Change()
    On Error GoTo errHandler
    Me.cmdOK.Enabled = (Len(txtReason) > 10)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtReason_Change", , EA_NORERAISE
    HandleError
End Sub

'Private Sub txtFlag_LostFocus()
'    If flgLoading Then Exit Sub
'    txtFlag = oProd.FlagText
'End Sub
'Private Sub txtFlag_Validate(Cancel As Boolean)
'    Cancel = mCancel
'End Sub
'Private Sub txtFlag_Change()
'Dim intPos As Integer
'    On Error Resume Next
'    mCancel = Not oProd.SetFlagtext(txtFlag)
'    If Err Then
'      Beep
'      intPos = txtFlag.SelStart
'      txtFlag = oProd.FlagText
'      txtFlag.SelStart = intPos - 1
'    End If
'End Sub

Private Sub txtRRP_GotFocus()
    On Error GoTo errHandler
    txtRRP = oProd.RRP
    AutoSelect txtRRP
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtRRP_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtRRP_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRRP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oProd.SetRRP(txtRRP) Then
        Cancel = True
    End If
    txtRRP = oProd.RRPF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtRRP_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtRRP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtUKPrice_GotFocus()
    On Error GoTo errHandler
    If oProd.SupplierCurrencyID > 0 Then
        If oProd.SupplierCurrencyID <> oPC.Configuration.Currencies.FindBySysname("GBP").ID And oProd.SupplierCurrencyID <> oPC.Configuration.DefaultCurrencyID Then
            MsgBox "The supplier of this product usually trades in " & oPC.Configuration.Currencies.FindCurrencyByID(oProd.SupplierCurrencyID).Description, vbOKOnly, "Warning"
        End If
    End If
    txtUKPrice = oProd.UKPrice
    AutoSelect txtUKPrice
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtUKPrice_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtUKPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtUKPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim lngRoundTo As Long
Dim tmpPrice As Long

    If flgLoading Then Exit Sub
'    If oProd.SupplierCurrencyID > 0 Then
'        If oProd.SupplierCurrencyID <> oPC.Configuration.Currencies.FindBySysname("GBP").ID Then
'            If MsgBox("The supplier of this product usually trades in " & oPC.Configuration.Currencies.FindCurrencyByID(oProd.SupplierCurrencyID).Description & vbCrLf & "Are you sure you want to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
'                txtUKPrice = oProd.UKPrice
'                Exit Sub
'            End If
'        End If
'    End If
    If Not oProd.SetUKPrice(txtUKPrice) Then
        Cancel = True
    End If
    txtUKPrice = oProd.UKPriceF
    If oProd.SupplierConversionToLocalFactor <> 0 And oProd.SupplierCurrencyID = oPC.Configuration.Currencies.FindBySysname("GBP").ID Then
        tmpPrice = oProd.UKPrice * oProd.SupplierConversionToLocalFactor
        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice)
        oProd.SP = (RoundUp(tmpPrice, lngRoundTo))
        Me.txtSP = oProd.SPF
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtUKPrice_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtUKPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtEUPrice_GotFocus()
    On Error GoTo errHandler
    If oProd.SupplierCurrencyID > 0 Then
        If oProd.SupplierCurrencyID <> oPC.Configuration.Currencies.FindBySysname("EUR").ID And oProd.SupplierCurrencyID <> oPC.Configuration.DefaultCurrencyID Then
            MsgBox "The supplier of this product usually trades in " & oPC.Configuration.Currencies.FindCurrencyByID(oProd.SupplierCurrencyID).Description, vbOKOnly, "Warning"
        End If
    End If
    txtEUPrice = oProd.EUPrice
    AutoSelect txtEUPrice
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtEUPrice_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEUPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEUPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim lngRoundTo As Long
Dim tmpPrice As Long
    If flgLoading Then Exit Sub
'    If oProd.SupplierCurrencyID > 0 Then
'        If oProd.SupplierCurrencyID <> oPC.Configuration.Currencies.FindBySysname("EUR").ID Then
'            If MsgBox("The supplier of this product usually trades in " & oPC.Configuration.Currencies.FindCurrencyByID(oProd.SupplierCurrencyID).Description & vbCrLf & "Are you sure you want to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
'                txtEUPrice = oProd.EUPrice
'                Exit Sub
'            End If
'        End If
'    End If
    If Not oProd.SetEUPrice(txtEUPrice) Then
        Cancel = True
    End If
    txtEUPrice = oProd.EUPriceF
    If oProd.SupplierConversionToLocalFactor <> 0 And oProd.SupplierCurrencyID = oPC.Configuration.Currencies.FindBySysname("EUR").ID Then
        tmpPrice = oProd.EUPrice * oProd.SupplierConversionToLocalFactor
        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice)
        oProd.SP = (RoundUp(tmpPrice, lngRoundTo))
        Me.txtSP = oProd.SPF
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtEUPrice_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEUPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub txtUSPrice_GotFocus()
    On Error GoTo errHandler
    If oProd.SupplierCurrencyID > 0 Then
        If oProd.SupplierCurrencyID <> oPC.Configuration.Currencies.FindBySysname("USD").ID And oProd.SupplierCurrencyID <> oPC.Configuration.DefaultCurrencyID Then
            MsgBox "The supplier of this product usually trades in " & oPC.Configuration.Currencies.FindCurrencyByID(oProd.SupplierCurrencyID).Description, vbOKOnly, "Warning"
        End If
    End If
    txtUSPrice = oProd.USPrice
    AutoSelect txtUSPrice
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtUSPrice_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtUSPrice_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtUSPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim lngRoundTo As Long
Dim tmpPrice As Long
    If flgLoading Then Exit Sub
'    If oProd.SupplierCurrencyID > 0 Then
'        If oProd.SupplierCurrencyID <> oPC.Configuration.Currencies.FindBySysname("USD").ID Then
'            If MsgBox("The supplier of this product usually trades in " & oPC.Configuration.Currencies.FindCurrencyByID(oProd.SupplierCurrencyID).Description & vbCrLf & "Are you sure you want to continue?", vbQuestion + vbYesNo, "Warning") = vbNo Then
'                txtUSPrice = oProd.USPrice
'                Exit Sub
'            End If
'        End If
'    End If
    If Not oProd.SetUSPrice(txtUSPrice) Then
        Cancel = True
    End If
    txtUSPrice = oProd.USPriceF
    If oProd.SupplierConversionToLocalFactor <> 0 And oProd.SupplierCurrencyID = oPC.Configuration.Currencies.FindBySysname("USD").ID Then
        tmpPrice = oProd.USPrice * oProd.SupplierConversionToLocalFactor
        lngRoundTo = oPC.Configuration.RoundingRules.GetRoundTo(tmpPrice)
        oProd.SP = (RoundUp(tmpPrice, lngRoundTo))
        Me.txtSP = oProd.SPF
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtUSPrice_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtUSPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtSSP_GotFocus()
    On Error GoTo errHandler
    txtSSP = oProd.SpecialPrice
    AutoSelect txtSSP
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtSSP_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSSP_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSSP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oProd.SetSpecialPrice(txtSSP) Then
        Cancel = True
    End If
    txtSSP = oProd.SpecialPriceF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtSSP_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSSP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtSP_GotFocus()
    On Error GoTo errHandler
    txtSP = oProd.SP
    AutoSelect txtSP
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtSP_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSP_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSP_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oProd.SetSP(txtSP) Then
        Cancel = True
    End If
    txtSP = oProd.SPF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtSP_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCost_GotFocus()
    On Error GoTo errHandler
    txtCost = oProd.Cost
    AutoSelect txtCost
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtCost_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtCost_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oProd.SetCost(txtCost) Then
        Cancel = True
    End If
    txtCost = oProd.CostF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtCost_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtCost_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
'Private Sub txtSpecialPrice_GotFocus()
'    txtSpecialPrice = oProd.SpecialPrice
'    AutoSelect txtSpecialPrice
'End Sub
'Private Sub txtSpecialPrice_Validate(Cancel As Boolean)
'    If flgLoading Then Exit Sub
'    If Not oProd.setspecialPrice(txtSpecialPrice) Then
'        Cancel = True
'    End If
'    txtSpecialPrice = oProd.SpecialPriceF
'End Sub

Private Sub txtSubtitle_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtSubtitle = oProd.SubTitle
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtSubtitle_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSubtitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSubtitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
   ' Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtSubtitle_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSubtitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtSubtitle_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetSubTitle(txtSubtitle)
    If Err Then
      Beep
      intPos = txtSubtitle.SelStart
      txtSubtitle = oProd.SubTitle
      txtSubtitle.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtSubtitle_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSubtitle_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtDescription_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtDescription = oProd.Description
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtDescription_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtDescription_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtDescription_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtDescription_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtDescription_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetDescription(txtDescription)
    If Err Then
      Beep
      intPos = txtDescription.SelStart
      txtDescription = oProd.Description
      txtDescription.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtDescription_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtDescription_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtComment_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtComment = oProd.Comment
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtComment_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtComment_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtComment_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtComment_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtComment_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtComment_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetComment(txtComment)
    If Err Then
      Beep
      intPos = txtComment.SelStart
      txtComment = oProd.Comment
      txtComment.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtComment_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtComment_Change", , EA_NORERAISE
    HandleError
End Sub




Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtTitle = oProd.Title
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
Dim intPos As Integer
    On Error GoTo errHandler
    On Error Resume Next
    mCancel = Not oProd.SetTitle(txtTitle)
    If Err Then
      Beep
      intPos = txtTitle.SelStart
      txtTitle = oProd.Title
      txtTitle.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtAuthor_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAuthor = oProd.Author
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtAuthor_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAuthor_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtAuthor_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtAuthor_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetAuthor(txtAuthor)
    If Err Then
      Beep
      intPos = txtAuthor.SelStart
      txtAuthor = oProd.Author
      txtAuthor.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtAuthor_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtAuthor_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPublisher = oProd.Publisher
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPublisher_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPublisher_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPublisher_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPublisher_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetPublisher(txtPublisher)
    If Err Then
      Beep
      intPos = txtPublisher.SelStart
      txtPublisher = oProd.Publisher
      txtPublisher.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPublisher_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPublisher_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubDate_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPubDate = oProd.PublicationDate
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPubDate_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPubDate_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubDate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPubDate_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPubDate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubDate_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetPublicationDate(txtPubDate)
    If Err Then
      Beep
      intPos = txtPubDate.SelStart
      txtPubDate = oProd.PublicationDate
      txtPubDate.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPubDate_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPubDate_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubPlace_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPubPlace = oProd.PublicationPlace
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPubPlace_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPubPlace_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubPlace_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPubPlace_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPubPlace_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubPlace_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetPublicationPlace(txtPubPlace)
    If Err Then
      Beep
      intPos = txtPubPlace.SelStart
      txtPubPlace = oProd.PublicationPlace
      txtPubPlace.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtPubPlace_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPubPlace_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtEdition = oProd.Edition
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtEdition_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEdition_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtEdition_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEdition_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetEdition(txtEdition)
    If Err Then
      Beep
      intPos = txtEdition.SelStart
      txtEdition = oProd.Edition
      txtEdition.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtEdition_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEdition_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oProd.Note
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtNote_LostFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetNote(txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oProd.Note
      txtNote.SelStart = intPos - 1
    End If
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtNote_Change", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtVAT_GotFocus()
    On Error GoTo errHandler
    txtVAT = oProd.VATRateF
    AutoSelect txtVAT
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtVAT_GotFocus", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtVAT_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtVAT_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oProd.SetVAT(txtVAT) Then
        Cancel = True
    End If
    txtVAT = oProd.VATRateF
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtVAT_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtVAT_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub LoadListView()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim i As Integer
Dim tmp As String

    lvwCE.ListItems.Clear
    For i = 1 To oProd.CatalogueEntries.Count
        Set objItm = Me.lvwCE.ListItems.Add
        With objItm
            .Key = oProd.CatalogueEntries(i).Key
            .text = oProd.CatalogueEntries(i).Serial & IIf(oProd.CatalogueEntries(i).IsDeleted, "(DEL)", "")
            .SubItems(1) = oProd.CatalogueEntries(i).PriceF
        End With
    Next i
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.LoadListView"
End Sub

