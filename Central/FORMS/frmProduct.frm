VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProduct 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Edit book"
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11430
   ControlBox      =   0   'False
   ForeColor       =   &H8000000D&
   Icon            =   "frmProduct.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleMode       =   0  'User
   ScaleWidth      =   15077.87
   Begin VB.CheckBox chkExSales 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Exclude from sales reporting"
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   6465
      TabIndex        =   65
      Top             =   270
      Width           =   2955
   End
   Begin VB.CommandButton cmdHelp 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   90
      Width           =   450
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
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   6090
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
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1650
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
      Left            =   10335
      Picture         =   "frmProduct.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5730
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
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
      Height          =   600
      Left            =   9405
      Picture         =   "frmProduct.frx":0894
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5745
      Width           =   930
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
      Top             =   900
      Width           =   2250
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   35
      Top             =   5775
      Width           =   4350
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3510
      Left            =   15
      TabIndex        =   10
      Top             =   2145
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   6191
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
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
      TabPicture(0)   =   "frmProduct.frx":0E1E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label32"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label31"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label18"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label26"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label17"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label12"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblSeesafe"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtUSPrice"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtUKPrice"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCost"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtSP"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtRRP"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkMAG"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "chkObsolete"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboCatHead"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cboSection"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboProductType"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdAddSection"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtSection"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProduct.frx":0E3A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSupplier"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblDeal"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label20"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label14"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtBinding"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdSetDefault"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtVAT"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdSupplier"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
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
         Left            =   -73425
         TabIndex        =   59
         Top             =   1665
         Width           =   4200
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
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   690
            Width           =   3870
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
            TabIndex        =   24
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
         Height          =   315
         Left            =   -64920
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   780
         Width           =   570
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
         Height          =   345
         Left            =   7290
         MultiLine       =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1005
         Width           =   3870
      End
      Begin VB.CommandButton cmdAddSection 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   315
         Left            =   6510
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1020
         Width           =   750
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
         Left            =   4335
         TabIndex        =   16
         Top             =   420
         Width           =   2490
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
         Left            =   -73440
         TabIndex        =   23
         Top             =   1125
         Width           =   1380
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
         Left            =   -71970
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1095
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
         Left            =   -73440
         TabIndex        =   22
         Top             =   720
         Width           =   1395
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
         Height          =   1155
         Left            =   9225
         TabIndex        =   21
         Top             =   2235
         Width           =   1800
         Begin VB.OptionButton optRP 
            Caption         =   "Reprinting"
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
            Left            =   255
            TabIndex        =   48
            Top             =   825
            Width           =   1335
         End
         Begin VB.OptionButton optOOP 
            Caption         =   "Out of print"
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
            Left            =   270
            TabIndex        =   47
            Top             =   547
            Width           =   1335
         End
         Begin VB.OptionButton optIP 
            Caption         =   "In print"
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
            Left            =   270
            TabIndex        =   46
            Top             =   255
            Value           =   -1  'True
            Width           =   1335
         End
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
         Left            =   3990
         TabIndex        =   17
         Top             =   1005
         Width           =   2490
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
         Left            =   3225
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1740
         Width           =   7785
      End
      Begin VB.CheckBox chkObsolete 
         Alignment       =   1  'Right Justify
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
         Left            =   7725
         TabIndex        =   20
         Top             =   3105
         Width           =   1245
      End
      Begin VB.CheckBox chkMAG 
         Alignment       =   1  'Right Justify
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
         Height          =   480
         Left            =   3195
         TabIndex        =   19
         Top             =   2340
         Width           =   1665
      End
      Begin VB.TextBox txtRRP 
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
         Left            =   1260
         TabIndex        =   11
         Top             =   735
         Width           =   1380
      End
      Begin VB.TextBox txtSP 
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
         Left            =   1260
         TabIndex        =   12
         Top             =   1110
         Width           =   1380
      End
      Begin VB.TextBox txtCost 
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
         Left            =   1260
         TabIndex        =   13
         Top             =   1485
         Width           =   1380
      End
      Begin VB.TextBox txtUKPrice 
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
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2250
         Width           =   1380
      End
      Begin VB.TextBox txtUSPrice 
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
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2625
         Width           =   1380
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
         Height          =   360
         Left            =   8325
         TabIndex        =   64
         Top             =   420
         Width           =   1650
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
         Left            =   7455
         TabIndex        =   63
         Top             =   480
         Width           =   810
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
         Left            =   -68970
         TabIndex        =   61
         Top             =   1410
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
         Left            =   -68970
         TabIndex        =   58
         Top             =   810
         Width           =   810
      End
      Begin VB.Label Label12 
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
         Left            =   3105
         TabIndex        =   52
         Top             =   495
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
         Left            =   -74670
         TabIndex        =   51
         Top             =   1155
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
         Left            =   -74670
         TabIndex        =   50
         Top             =   750
         Width           =   1080
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
         Left            =   3045
         TabIndex        =   45
         Top             =   1485
         Width           =   1755
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Section"
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
         Left            =   2730
         TabIndex        =   44
         Top             =   1020
         Width           =   1080
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
         Left            =   450
         TabIndex        =   43
         Top             =   750
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
         Left            =   450
         TabIndex        =   42
         Top             =   1125
         Width           =   750
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
         Left            =   450
         TabIndex        =   41
         Top             =   1485
         Width           =   750
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
         Left            =   330
         TabIndex        =   40
         Top             =   2295
         Width           =   870
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
         Left            =   330
         TabIndex        =   39
         Top             =   2670
         Width           =   870
      End
      Begin VB.Label lblDeal 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   -68100
         TabIndex        =   38
         Top             =   1335
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
         Left            =   -68085
         TabIndex        =   37
         Top             =   750
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
      Top             =   660
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
      Top             =   1035
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
      Height          =   330
      Left            =   6480
      TabIndex        =   8
      Top             =   1785
      Width           =   2520
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
      TabIndex        =   7
      Top             =   1410
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
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   510
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
      Left            =   750
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "E.A.N."
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
      Left            =   2490
      TabIndex        =   55
      Top             =   150
      Width           =   480
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
      TabIndex        =   36
      Top             =   660
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
      TabIndex        =   34
      Top             =   675
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
      TabIndex        =   33
      Top             =   1065
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
      Left            =   5775
      TabIndex        =   32
      Top             =   1845
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
      TabIndex        =   31
      Top             =   1455
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN"
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
      Left            =   0
      TabIndex        =   27
      Top             =   165
      Width           =   660
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


Private Sub chkExSales_Click()
    oProd.ExcludeFromSales = IIf(Me.chkExSales = 1, True, False)
End Sub

Private Sub cmdChangeType_Click()
    On Error GoTo errHandler
    If MsgBox("You want to change this product to be a general product (non book)?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    Else
        oProd.SetProductType "G"
        oProd.ApplyEdit
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdChangeType_Click", , EA_NORERAISE
    HandleError
End Sub

Sub Component(pProduct As a_Product, Optional pPrevForm As Form)
    On Error GoTo errHandler
    Set frmPrevious = pPrevForm
    Set oProd = pProduct
    oProd.BeginEdit
    oProd.SetBook
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.Component(pProduct,pPrevForm)", Array(pProduct, pPrevForm)
End Sub

Sub cmdHelp_Click()
    On Error GoTo errHandler
Dim frm As New frmHelpGen
Dim tmp As String
    tmp = LoadResString(101)
    frm.Component tmp, "EAN and ISBN codes", 6000, 2000
    frm.Show vbModal
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdHelp_Click", , EA_NORERAISE
    HandleError
End Sub
'Private Sub cboSection_Click()
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    oProd.SetSection cboSection
'    txtSection = oProd.Section
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cboSection_Click", , EA_NORERAISE
'    HandleError
'End Sub
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

'Private Sub cboCatHead_Click()
'    On Error GoTo errHandler
'    If flgLoading Then Exit Sub
'    oProd.SetCatalogueheadingID oPC.Configuration.CatalogueHeadings.Key(cboCatHead)
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cboCatHead_Click", , EA_NORERAISE
'    HandleError
'End Sub

'Private Sub chkNonStock_Click()
'    If chkNonStock Then
'        oProd.SetNONStock
'    Else
'        oProd.SetBook
'    End If
'End Sub

Private Sub chkObsolete_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.Obsolete = chkObsolete
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.chkObsolete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkMAG_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetMagsEtc = chkMAG
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.chkMAG_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdAddCopy_Click()
'    On Error GoTo errHandler
'Dim frm As frmCopy
'Dim oCopy As a_Copy
'Dim tmpCopy As a_Copy
'    Set frm = New frmCopy
'    Set oCopy = oProd.Copies.Add
'
'    If Grid1.Bookmark > 0 Then
'        Set tmpCopy = oProd.Copies(Grid1.Bookmark)
'    Else
'        Set tmpCopy = Nothing
'    End If
'
'
'    frm.Component oCopy, tmpCopy
'    frm.Show vbModal
'    Set oCopy = Nothing
'    Set frm = Nothing
'    LoadCopies
'    Grid1.ReBind
'    Grid1.ReBind
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdAddCopy_Click", , EA_NORERAISE
'    HandleError
'End Sub


Private Sub cmdDelete_Click()
    On Error GoTo errHandler

    If MsgBox("Confirm deletion of product: " & oProd.Title & vbCrLf & "This is permanent!", vbOKCancel + vbExclamation, "Confirm") = vbOK Then
        oProd.Delete
        oProd.ApplyEdit
        Unload Me
    End If


    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdAddSection_Click()
'    On Error GoTo errHandler
'    oProd.SetSection cboSection
'    txtSection = oProd.Section
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdAddSection_Click", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub cmdGenerateEAN_Click()
'    On Error GoTo errHandler
'Dim oProdCode As New z_ProdCode
'    oProdCode.SetCodesForBook txtCode
'    oProd.SetEAN oProdCode.EAN
'    txtEAN = oProd.EAN
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdGenerateEAN_Click", , EA_NORERAISE
'    HandleError
'End Sub

'Private Sub cmdNewWant_Click()
'    On Error GoTo errHandler
'Dim frm As frmWant
''Dim oWant As a_Want
'   ' Set oWant = oProd.Wants.Add
'    Set frm = New frmWant
'    frm.Component oProd
'    frm.Show vbModal
'    LoadWants
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdNewWant_Click", , EA_NORERAISE
'    HandleError
'End Sub
'
'Private Sub cmdRemove_Click()
'    On Error GoTo errHandler
'Dim oCopy As a_Copy
'    Set oCopy = oProd.Copies(XA(Grid1.Bookmark, 6))
'    oCopy.BeginEdit
'    oCopy.Delete
'    oCopy.ApplyEdit
'    LoadCopies
'    Grid1.ReBind
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdRemove_Click", , EA_NORERAISE
'    HandleError
'End Sub
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.UPdateFromXArray"
End Sub


Private Sub cmdSetDefault_Click()
    On Error GoTo errHandler
    Me.txtVAT = oPC.Configuration.vatRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdSetDefault_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oProd.IsEditing Then oProd.CancelEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub



Private Sub oProd_Valid(strMsg As String)
    On Error GoTo errHandler
    Me.txtErrors = strMsg
    Me.cmdOK.Enabled = (strMsg = "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.oProd_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oProd.CancelEdit
    Unload Me
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
    WaitMsg "Saving product . . .", True, Me
    oProd.SetBook
    oProd.ApplyEdit lngResult, strMsg
    If lngResult = 99 Then
        WaitMsg "", False, Me
        If strMsg = "TIMEOUT" Then
            MsgBox "The operation has timed out. The record is probably locked by another user." & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
        ElseIf strMsg = "DUPLICATE" Then
            MsgBox "Invalid values - check that the code is has not been already used", vbInformation, "Save failed"
        End If
    Else
        If frmPrevious Is Nothing Then
            Set frmPreview = New frmProductPrev
        Else
            Set frmPreview = frmPrevious
        End If
        frmPreview.Component oProd
        frmPreview.RefreshForm
        frmPreview.Show
        WaitMsg "", False, Me
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    Left = 10
    Top = 10
    Width = 11500
    Height = 6800
    LoadControls
    Me.SSTab1.Tab = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtCode = oProd.code
    txtEAN = oProd.EAN
    txtTitle = oProd.Title
    txtSubtitle = oProd.Subtitle
    txtAuthor = oProd.Author
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    txtPubPlace = oProd.PublicationPlace
    Me.txtPubDate = oProd.PublicationDate
    Me.txtBinding = oProd.BindingCode
    txtRRP = oProd.RRPF
    txtSP = oProd.SPF
    txtCost = oProd.CostF
    txtBIC = oProd.BIC
    txtSection = oProd.Section
    Me.txtNote = oProd.Note
    Me.txtVAT = oProd.vatratef
    LoadCombo cboSection, oPC.Configuration.Sections
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    cboProductType = oPC.Configuration.ProductTypes.Item(oProd.ProductTypeID)
    Me.chkMAG = IIf(oProd.IsMagsEtc, 1, 0)
    Me.chkObsolete = IIf(oProd.Obsolete, 1, 0)
    Me.lblSupplier.Caption = oProd.lastsuppliername
    Me.lblDeal.Caption = oProd.lastDealDescription
    Select Case oProd.Status
    Case "O"
        optOOP.Value = True
    Case "R"
        optRP.Value = True
    Case Else
        optIP.Value = True
    End Select
    Me.chkExSales = IIf(oProd.ExcludeFromSales, 1, 0)
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.LoadControls"
End Sub

Private Sub optIP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enInPrint
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.optIP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOOP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enOutOfPrint
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.optOOP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optRP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enAwaitingReprint
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

Private Sub txtBinding_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtBinding = oProd.BindingCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtBinding_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBinding_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtBinding_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetCode(txtCode)
    If Err Then
      Beep
      intPos = txtCode.SelStart
      txtCode = oProd.code
      txtCode.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtCode_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not oProd.SetCode(txtCode)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtEAN_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.SetEAN(txtEAN)
    If Err Then
      Beep
      intPos = txtEAN.SelStart
      txtEAN = oProd.EAN
      txtEAN.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEAN_Change", , EA_NORERAISE
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
    Cancel = Not oProd.SetEAN(txtEAN)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEAN_Validate(Cancel)", Cancel, EA_NORERAISE
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtRRP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtSection_Validate(Cancel As Boolean)
'    On Error GoTo errHandler
'    oProd.SetSectionAll txtSection
'    txtSection = oProd.Section
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.txtSection_Validate(Cancel)", Cancel, EA_NORERAISE
'    HandleError
'End Sub

Private Sub txtSP_GotFocus()
    On Error GoTo errHandler
    txtSP = oProd.SP
    AutoSelect txtSP
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
    txtSubtitle = oProd.Subtitle
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSubtitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSubtitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
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
      txtSubtitle = oProd.Subtitle
      txtSubtitle.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtSubtitle_Change", , EA_NORERAISE
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
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Change()
    On Error GoTo errHandler
Dim intPos As Integer
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
    ErrorIn "frmProduct.txtTitle_Change", , EA_NORERAISE
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPublisher_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPubDate_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubDate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtPubPlace_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubPlace_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtEdition_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    mCancel = Not oProd.setnote(txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oProd.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProduct.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

