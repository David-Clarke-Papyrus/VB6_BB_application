VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmProductNBPrev 
   BackColor       =   &H00D3D3CB&
   Caption         =   "General stock"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11820
   ControlBox      =   0   'False
   Icon            =   "frmProductNBPrev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleMode       =   0  'User
   ScaleWidth      =   15592.34
   Begin VB.CommandButton cmdForward 
      Height          =   390
      Left            =   4095
      Picture         =   "frmProductNBPrev.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   5970
      Width           =   435
   End
   Begin VB.CommandButton cmdBack 
      Height          =   390
      Left            =   3585
      Picture         =   "frmProductNBPrev.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   5970
      Width           =   435
   End
   Begin TrueOleDBGrid60.TDBGrid StGrid 
      Height          =   1485
      Left            =   9825
      OleObjectBlob   =   "frmProductNBPrev.frx":0A1E
      TabIndex        =   72
      Top             =   45
      Width           =   1650
   End
   Begin VB.TextBox txtLastOrdered 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6450
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   195
      Width           =   960
   End
   Begin VB.TextBox txtLastOrderedQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   195
      Width           =   555
   End
   Begin VB.TextBox txtLastOrderedPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   195
      Width           =   885
   End
   Begin VB.TextBox txtLastReceivedPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   510
      Width           =   885
   End
   Begin VB.TextBox txtLastReceivedQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   510
      Width           =   555
   End
   Begin VB.TextBox txtLastReceived 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6450
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   510
      Width           =   960
   End
   Begin VB.TextBox txtLastSoldPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8010
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   825
      Width           =   885
   End
   Begin VB.TextBox txtLastSoldQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   825
      Width           =   555
   End
   Begin VB.TextBox txtLastSoldDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6450
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   825
      Width           =   960
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
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   1170
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
      Left            =   10320
      Picture         =   "frmProductNBPrev.frx":4CFD
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5925
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
      Left            =   9210
      Picture         =   "frmProductNBPrev.frx":4DA8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5925
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Find By ISBN"
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   225
      TabIndex        =   8
      Top             =   5880
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   270
         Width           =   1995
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4005
      Left            =   210
      TabIndex        =   4
      Top             =   1875
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   6
      TabHeight       =   535
      BackColor       =   13882315
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&1. Stock"
      TabPicture(0)   =   "frmProductNBPrev.frx":50B2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label21"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblStatus"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label40"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label29"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label19"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label18"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label16"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "MMGRID"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "APPGRID"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "COGrid"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "POGrid"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtOnHand"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtReserved"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtTotalSold"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtTotalOSPO"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtTotalOSCO"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtTotalOSAP"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdALLMM"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DTMMSince"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtReturnable"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdRecon"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtSSP"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtCost"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtSP"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtRRP"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProductNBPrev.frx":50CE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbPSECs"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtProductType"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtDefaultDeliveryDays"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtSS"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtPublisher"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtEdition"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtCategory"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtVAT"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label35"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label34"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label33"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label23"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label7"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label8"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblSupplier"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblObsolete"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label26"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label10"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "&3. Statistics"
      TabPicture(2)   =   "frmProductNBPrev.frx":50EA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtDateAdded"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "txtDateLastModified"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "txtLastCounted"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtLastCountedQty"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtLastCountedPrice"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label12"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label22"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label1"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "4. Substitute products"
      TabPicture(3)   =   "frmProductNBPrev.frx":5106
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label43"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label44"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "GSF"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "GSB"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cmdSub"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cmdSubs"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
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
         Left            =   -74895
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   390
         Width           =   1920
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
         Left            =   -65295
         MaskColor       =   &H00C4BCA4&
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   480
         Width           =   1305
      End
      Begin VB.TextBox txtRRP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   1635
         Width           =   1170
      End
      Begin VB.TextBox txtSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   1635
         Width           =   1200
      End
      Begin VB.TextBox txtCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2235
         Width           =   1170
      End
      Begin VB.TextBox txtSSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1725
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   2235
         Width           =   1200
      End
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
         Left            =   7155
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   3525
         Width           =   720
      End
      Begin VB.ListBox lbPSECs 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   225
         Left            =   -66870
         TabIndex        =   60
         Top             =   1140
         Width           =   2715
      End
      Begin VB.TextBox txtDateAdded 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -66045
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   3060
         Width           =   2160
      End
      Begin VB.TextBox txtDateLastModified 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -66045
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   3480
         Width           =   2160
      End
      Begin VB.TextBox txtLastCounted 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2685
         Width           =   1380
      End
      Begin VB.TextBox txtLastCountedQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -71655
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2685
         Width           =   555
      End
      Begin VB.TextBox txtLastCountedPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -71040
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2685
         Width           =   960
      End
      Begin VB.Frame Frame2 
         Caption         =   "Aged stock"
         ForeColor       =   &H8000000D&
         Height          =   810
         Left            =   -74565
         TabIndex        =   46
         Top             =   570
         Width           =   5220
         Begin VB.TextBox txtAgedDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   330
            Width           =   1170
         End
         Begin VB.TextBox txt6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2415
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt18Plus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00DBFAFB&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   330
            Width           =   825
         End
      End
      Begin VB.TextBox txtReturnable 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   3270
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txtProductType 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -66870
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   630
         Width           =   2700
      End
      Begin MSComCtl2.DTPicker DTMMSince 
         Height          =   360
         Left            =   9675
         TabIndex        =   40
         Top             =   3510
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   221839361
         CurrentDate     =   37656
         MaxDate         =   55153
         MinDate         =   34820
      End
      Begin VB.CommandButton cmdALLMM 
         BackColor       =   &H00C4BCA4&
         Caption         =   "All movements since:"
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3510
         Width           =   1695
      End
      Begin VB.TextBox txtDefaultDeliveryDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   3360
         Width           =   1395
      End
      Begin VB.TextBox txtSS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2515
         Width           =   1395
      End
      Begin VB.TextBox txtTotalOSAP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         TabIndex        =   30
         Top             =   1680
         Width           =   3705
      End
      Begin VB.TextBox txtTotalOSCO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3510
         Width           =   3705
      End
      Begin VB.TextBox txtPublisher 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1255
         Width           =   3990
      End
      Begin VB.TextBox txtEdition 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1675
         Width           =   3990
      End
      Begin VB.TextBox txtTotalOSPO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1650
         Width           =   3705
      End
      Begin VB.TextBox txtTotalSold 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   2130
         TabIndex        =   15
         Top             =   825
         Width           =   750
      End
      Begin VB.TextBox txtReserved 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   825
         Width           =   750
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   825
         Width           =   750
      End
      Begin VB.TextBox txtCategory 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   840
         Width           =   2370
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2095
         Width           =   1395
      End
      Begin TrueOleDBGrid60.TDBGrid POGrid 
         Height          =   1230
         Left            =   3075
         OleObjectBlob   =   "frmProductNBPrev.frx":5122
         TabIndex        =   22
         Top             =   420
         Width           =   3960
      End
      Begin TrueOleDBGrid60.TDBGrid COGrid 
         Height          =   1245
         Left            =   3090
         OleObjectBlob   =   "frmProductNBPrev.frx":9B11
         TabIndex        =   29
         Top             =   2235
         Width           =   3945
      End
      Begin TrueOleDBGrid60.TDBGrid APPGRID 
         Height          =   1245
         Left            =   7110
         OleObjectBlob   =   "frmProductNBPrev.frx":E01C
         TabIndex        =   31
         Top             =   435
         Width           =   4065
      End
      Begin TrueOleDBGrid60.TDBGrid MMGRID 
         Height          =   1215
         Left            =   7155
         OleObjectBlob   =   "frmProductNBPrev.frx":12528
         TabIndex        =   32
         Top             =   2250
         Width           =   4035
      End
      Begin TrueOleDBGrid60.TDBGrid GSB 
         Height          =   1320
         Left            =   -73485
         OleObjectBlob   =   "frmProductNBPrev.frx":1666F
         TabIndex        =   92
         Top             =   2490
         Width           =   9495
      End
      Begin TrueOleDBGrid60.TDBGrid GSF 
         Height          =   1500
         Left            =   -73485
         OleObjectBlob   =   "frmProductNBPrev.frx":1A6DB
         TabIndex        =   93
         Top             =   900
         Width           =   9495
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "This item can substitute for these . . ."
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   -74835
         TabIndex        =   95
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Possible substitutions for this item"
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   -74835
         TabIndex        =   94
         Top             =   2910
         Width           =   1185
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R.R.P."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   270
         TabIndex        =   71
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sell.P."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1740
         TabIndex        =   70
         Top             =   1410
         Width           =   540
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cost (avg.)"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   150
         TabIndex        =   69
         Top             =   1995
         Width           =   1110
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Special"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1665
         TabIndex        =   68
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sections"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -68130
         TabIndex        =   61
         Top             =   1140
         Width           =   1080
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Added"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -67440
         TabIndex        =   59
         Top             =   3135
         Width           =   1290
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Last modified"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -67530
         TabIndex        =   58
         Top             =   3525
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Last counted"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74625
         TabIndex        =   57
         Top             =   2715
         Width           =   1395
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Returnable"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   210
         TabIndex        =   45
         Top             =   3315
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Product type"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -68130
         TabIndex        =   43
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
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
         Left            =   1230
         TabIndex        =   41
         Top             =   3645
         Width           =   1590
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Lead time"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74775
         TabIndex        =   38
         Top             =   3390
         Width           =   1590
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Last supplied by"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   -74775
         TabIndex        =   36
         Top             =   2992
         Width           =   1590
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Order by seesafe"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74775
         TabIndex        =   34
         Top             =   2555
         Width           =   1590
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturer"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74385
         TabIndex        =   27
         Top             =   1305
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -73830
         TabIndex        =   26
         Top             =   1711
         Width           =   645
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -73080
         TabIndex        =   19
         Top             =   2940
         Width           =   3135
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Total sold"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2010
         TabIndex        =   18
         Top             =   570
         Width           =   870
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Reserved"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1020
         TabIndex        =   17
         Top             =   570
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "On hand"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   195
         TabIndex        =   16
         Top             =   570
         Width           =   705
      End
      Begin VB.Label lblObsolete 
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -72600
         TabIndex        =   12
         Top             =   2935
         Width           =   1380
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Category"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74265
         TabIndex        =   7
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "V.A.T. Rate"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74265
         TabIndex        =   6
         Top             =   2118
         Width           =   1080
      End
   End
   Begin VB.TextBox txtSubtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   645
      Left            =   1155
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   870
      Width           =   4395
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   660
      Left            =   1155
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   4395
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last ord'd"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5445
      TabIndex        =   89
      Top             =   225
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last rec'd"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5340
      TabIndex        =   88
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   6495
      TabIndex        =   87
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   7290
      TabIndex        =   86
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   8025
      TabIndex        =   85
      Top             =   0
      Width           =   840
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last sold"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5535
      TabIndex        =   84
      Top             =   840
      Width           =   870
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "(RRP)"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8895
      TabIndex        =   83
      Top             =   285
      Width           =   765
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "(Cost)"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8895
      TabIndex        =   82
      Top             =   570
      Width           =   765
   End
   Begin VB.Label lblServiceITem 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "NON STOCK-TAKE ITEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   240
      TabIndex        =   35
      Top             =   1590
      Width           =   2460
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   405
      TabIndex        =   3
      Top             =   900
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   195
      Width           =   945
   End
End
Attribute VB_Name = "frmProductNBPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngID As Long
'Private lslist As ListItem

Dim XSF As XArrayDB 'substitutes for these
Dim XSB As XArrayDB 'can be substituted by

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
Sub component(pProduct As a_Product, Optional pstrTime As String)
    On Error GoTo errHandler
    strTime = pstrTime
strTime = strTime & "Start frmProductPrev:component:" & Now() & vbCrLf
    Set oProd = Nothing
    Set oProd = pProduct

'    Set tlCatHead = Nothing
'    Set tlCatHead = New z_TextList
'    tlCatHead.Load ltCatalogueHeadings
strTime = strTime & "End frmProductPrev:component:" & Now() & vbCrLf
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.component(pProduct,pstrTime)", Array(pProduct, pstrTime)
End Sub

Private Sub cmdExpand_Click()
    On Error GoTo errHandler
Dim frm As frmSalesCH
    Set frm = New frmSalesCH
    frm.component oProd
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdExpand_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdBack_Click()
    On Error GoTo errHandler
Dim strPID As String
    If Forms(0).frmBrowseProd Is Nothing Then Exit Sub
    strPID = Forms(0).frmBrowseProd.PrevPID
    If strPID > "" Then
        Set oProd = Nothing
        Set oProd = New a_Product
        oProd.Load strPID, 0
        LoadControls
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdBack_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdForward_Click()
    On Error GoTo errHandler
Dim strPID As String
    If Forms(0).frmBrowseProd Is Nothing Then Exit Sub
    strPID = Forms(0).frmBrowseProd.NextPID
    If strPID > "" Then
        Set oProd = Nothing
        Set oProd = New a_Product
        oProd.Load strPID, 0
        LoadControls
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdForward_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRecon_Click()
    On Error GoTo errHandler
Dim frmMM As frmMovements
    oProd.ReloadRecentMovements
    Set frmMM = New frmMovements
    frmMM.component oProd, Me.CurrentY, Me.CurrentX
    frmMM.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdRecon_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSales_Click()
    On Error GoTo errHandler
Dim frm As frmSalesCH
    Set frm = New frmSalesCH
    frm.component oProd
    frm.Show

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdSales_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSub_Click()
    On Error GoTo errHandler
Dim frm As New frmSubstitute
    frm.Component2 Trim(oProd.PID)
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdSub_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSubs_Click()
    On Error GoTo errHandler
    Set XSF = New XArrayDB
    Set XSB = New XArrayDB
    oProd.GetSubstitutes XSF, XSB
    Set GSF.Array = XSF
    Set GSB.Array = XSB
    GSF.ReBind
    GSB.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdSubs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuAdjust.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.SetMenu"
End Sub

Private Sub cboCatHead_Click()
    On Error GoTo errHandler
'    oProd.setCatalogueheadingID tlCatHead.Key(cboCatHead)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cboCatHead_Click", , EA_NORERAISE
    HandleError
End Sub

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
    ErrorIn "frmProductNBPrev.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub APPGRID_DblClick()
    On Error GoTo errHandler
Dim frm As New frmAPPPreview
    If IsNull(APPGRID.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    frm.component XE(APPGRID.Bookmark, 6)
    frm.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProductNBPrev: APPGRID_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmProductNBPrev: APPGRID_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.APPGRID_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_AxisLabelSelected(axisID As Integer, AxisIndex As Integer, labelSetIndex As Integer, LabelIndex As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.chart1_AxisLabelSelected(axisID,AxisIndex,labelSetIndex,LabelIndex," & _
        "MouseFlags,Cancel)", Array(axisID, AxisIndex, labelSetIndex, LabelIndex, MouseFlags, Cancel), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_AxisTitleSelected(axisID As Integer, AxisIndex As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.chart1_AxisTitleSelected(axisID,AxisIndex,MouseFlags,Cancel)", _
         Array(axisID, AxisIndex, MouseFlags, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_ChartSelected(MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.chart1_ChartSelected(MouseFlags,Cancel)", Array(MouseFlags, Cancel), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_FootnoteSelected(MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.chart1_FootnoteSelected(MouseFlags,Cancel)", Array(MouseFlags, Cancel), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_LegendSelected(MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.chart1_LegendSelected(MouseFlags,Cancel)", Array(MouseFlags, Cancel), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_PlotSelected(MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.chart1_PlotSelected(MouseFlags,Cancel)", Array(MouseFlags, Cancel), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
'MsgBox "Series " & Series
'MsgBox "Datapoint " & DataPoint
    Select Case Series
    Case 1
        MsgBox "Sales qty = " & oProd.CurrentSales.FindByWeek(DataPoint).Qty & vbCrLf & "Sales value = " & oProd.CurrentSales.FindByWeek(DataPoint).ValuF
    Case 2
   '     MsgBox "Sales qty = " & oProd.PreviousSales.FindByWeek(DataPoint).Qty & vbCrLf & "Sales value = " & oProd.PreviousSales.FindByWeek(DataPoint).ValuF
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.chart1_PointSelected(Series,DataPoint,MouseFlags,Cancel)", Array(Series, _
         DataPoint, MouseFlags, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
'Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.chart1_SeriesSelected(Series,MouseFlags,Cancel)", Array(Series, _
         MouseFlags, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdALLMM_Click()
    On Error GoTo errHandler
Dim frmMM As frmMovements
    oProd.ReloadMovements Me.DTMMSince
    Set frmMM = New frmMovements
    frmMM.component oProd, Me.CurrentY, Me.CurrentX
    frmMM.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdALLMM_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmProductNB
    Set frm = New frmProductNB
    frm.component oProd, Me
    frm.Show
    Exit Sub
Errh:
    MsgBox Error

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cmdsearchisbn_Click()
    On Error GoTo errHandler
    Set oProd = Nothing
    Set oProd = New a_Product
    With oProd
    .Load "", 0, txtisbnsearch
       
    Me.Caption = "Stock code: " & .code
    txtSubtitle = .SubTitle
    txtTitle = .Title
    txtPublisher = .Publisher
        
    End With
    LoadControls
    LoadStock
    LoadMovements
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.cmdsearchisbn_Click", , EA_NORERAISE
    HandleError
End Sub




Private Sub COGrid_DblClick()
    On Error GoTo errHandler
Dim frm As New frmCOPreview
    If IsNull(COGrid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    frm.component XD(COGrid.Bookmark, 6), False
    frm.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProductNBPrev: COGrid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmProductNBPrev: COGrid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.COGrid_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub Grid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo errHandler
'Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.Grid1_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, _
         KeyAscii, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
            Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub




Private Sub MMGRID_DblClick()
    On Error GoTo errHandler
Dim strType As String
Dim frm As Form
Dim i As Integer


    If IsNull(MMGRID.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass

    strType = XF(MMGRID.Bookmark, 4)
    Select Case strType
    Case "APP"
        Set frm = New frmAPPPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "APR"
        Set frm = New frmAPPRPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "INV"
        Set frm = New frmInvoicePreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "GDN"
        Set frm = New frmGDNPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "CS"
        Set frm = New frmCSPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "TF"
        Set frm = New frmTFPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "DEL"
        Set frm = New frmDELPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
    Case "CN"
        Set frm = New frmCNPreview
        frm.component XF(MMGRID.Bookmark, 6)
        frm.Show
'    Case "RET"
'        Set frm = New frmReturn3
'        Set oR = new d_R
'
'        frm.Component XF(MMGRID.Bookmark, 6), ""
'        frm.Show
    End Select
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProductNBPrev: MMGRID_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmProductNBPrev: MMGRID_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.MMGRID_DblClick", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuAdjust()
    On Error GoTo errHandler
Dim frm As New frmStockAdjust
    frm.component oProd
    frm.Show vbModal
    If frm.Cancelled = False Then
        
        Me.txtLastCounted = Now
        Me.txtLastCountedQty = frm.Counted
        Me.txtOnHand = frm.Count
    End If
    Unload frm
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.mnuAdjust"
End Sub

Private Sub POGrid_DblClick()
    On Error GoTo errHandler
Dim frm As New frmPOPreview
    If IsNull(POGrid.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass
    frm.component XC(POGrid.Bookmark, 6)
    frm.Show
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProductNBPrev: POGrid_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmProductNBPrev: POGrid_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.POGrid_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub StGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.StGrid_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11700
        Height = 7400
    End If
strTime = strTime & "Start frmProductPrev:Load:" & Now() & vbCrLf
    LoadControls
strTime = strTime & "End frmProductPrev:Load:" & Now() & vbCrLf
    Me.SSTab1.Tab = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Public Sub RefreshForm()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.RefreshForm"
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    Me.Caption = "Stock code: " & oProd.code & "      EAN: " & oProd.EAN
    txtTitle = oProd.Title
    txtSubtitle = oProd.SubTitle
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    Me.txtCategory = oPC.Configuration.Sections.Item(oProd.CategoryID)
    Me.txtProductType = oPC.Configuration.ProductTypes.Item(oProd.ProductTypeID)
    txtDefaultDeliveryDays = oProd.DefaultDeliveryDays
    Me.txtRRP = oProd.RRPF
    Me.txtSP = oProd.SPF
    Me.txtCost = oProd.CostF
    Me.txtTotalSold = oProd.QtyTotalSold
    Me.txtAgedDate = ""
    Me.txtCost = oProd.CostF
    Me.txtDateAdded = oProd.DateRecordAddedF
    Me.txtDateLastModified = oProd.DateLastModifiedF
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
    Me.txtVAT = oProd.VATRateToUseF
    Me.txtReserved = oProd.QtyReservedF
    Me.txtReturnable = oProd.ReturnAvailability
    Me.lblServiceITem.Visible = oProd.IsServiceItem
    Me.lblObsolete = IIf(oProd.Obsolete, "obsolete", "")
    txtSS = IIf(oProd.Seesafe = 1, "Yes", "")
    Me.lblSupplier.Caption = oProd.LastSupplierName
    lblStatus = oProd.StatusF
    DTMMSince.Value = IIf(oProd.DateLastCounted < DTMMSince.MinDate, DTMMSince.MinDate, oProd.DateLastCounted)
    LoadMovements

    LoadStock
    LoadProductSections

    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.LoadControls"
End Sub
Private Sub LoadProductSections()
    On Error GoTo errHandler
Dim oPSEC As a_ProductSection
    With Me.lbPSECs
        .Clear
        For Each oPSEC In oProd.ProductSections
            .AddItem oPSEC.Description & "  " & oPSEC.PriorityF
        Next
        If .ListCount > 0 Then .ListIndex = 0
    End With

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.LoadProductSections"
End Sub

Private Sub LoadStock()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XB = New XArrayDB
    XB.Clear
    XB.ReDim 1, oProd.Stores.Count, 1, 6
    For lngIndex = 1 To oProd.Stores.Count
        XB.Value(lngIndex, 1) = oProd.Stores(lngIndex).StoreName
        XB.Value(lngIndex, 2) = oProd.Stores(lngIndex).QtyOnHand
        XB.Value(lngIndex, 3) = oProd.Stores(lngIndex).QtyReserved
        XB.Value(lngIndex, 4) = oProd.Stores(lngIndex).QtyOnBackorder
        XB.Value(lngIndex, 5) = oProd.Stores(lngIndex).QtyonOrder
        XB.Value(lngIndex, 6) = oProd.Stores(lngIndex).QtyCopiesOnHand
    Next
    XB.QuickSort 1, oProd.Stores.Count, 1, XORDER_ASCEND, XTYPE_STRING
    StGrid.Array = XB
    StGrid.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.LoadStock"
End Sub

Private Sub LoadMovements()
    On Error GoTo errHandler
Dim lngQtyOutstanding As Long
Dim lngQtySpecial As Long
Dim lngQtyApps As Long

Dim oOSPOD As d_OSSOrder
Dim oOSCOD As d_OSCOrder
Dim oOSAPP As d_OSAPPRO
Dim oMM As d_MM

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
    MMGRID.Splits(0).Caption = "Movements since last count (" & oProd.dateLastCountedF & ":" & oProd.QtyLastCountedF & ")"
    
    txtTotalOSPO = "We are awaiting " & lngQtyOutstanding & " copies."
    txtTotalOSCO = "Customers are awaiting " & lngQtySpecial & " copies."
    txtTotalOSAP = "Expecting return of " & lngQtyApps & " copies."
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.LoadMovements"
End Sub
Private Sub LoadPOs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XC = New XArrayDB
    XC.Clear
    XC.ReDim 1, oProd.OSPOs.Count, 1, 7
    For lngIndex = 1 To oProd.OSPOs.Count
        XC.Value(lngIndex, 1) = oProd.OSPOs(lngIndex).DOCCode
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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.LoadPOs"
End Sub
Private Sub LoadCOs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XD = New XArrayDB
    XD.Clear
    XD.ReDim 1, oProd.OSCOs.Count, 1, 7
    For lngIndex = 1 To oProd.OSCOs.Count
        XD.Value(lngIndex, 1) = oProd.OSCOs(lngIndex).DOCCode
        XD.Value(lngIndex, 2) = oProd.OSCOs(lngIndex).DocDateF
        XD.Value(lngIndex, 3) = oProd.OSCOs(lngIndex).TPNAME
        XD.Value(lngIndex, 4) = oProd.OSCOs(lngIndex).COLQty
        XD.Value(lngIndex, 5) = oProd.OSCOs(lngIndex).COLCollected
        XD.Value(lngIndex, 6) = oProd.OSCOs(lngIndex).TRID
        XD.Value(lngIndex, 7) = oProd.OSCOs(lngIndex).DateForSort
    Next
    XD.QuickSort 1, oProd.OSCOs.Count, 7, XORDER_DESCEND, XTYPE_STRING
    COGrid.Array = XD
    COGrid.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.LoadCOs"
End Sub
Private Sub LoadAPs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XE = New XArrayDB
    XE.Clear
    XE.ReDim 1, oProd.OSAPs.Count, 1, 7
    For lngIndex = 1 To oProd.OSAPs.Count
        XE.Value(lngIndex, 1) = oProd.OSAPs(lngIndex).DOCCode
        XE.Value(lngIndex, 2) = oProd.OSAPs(lngIndex).DocDateF
        XE.Value(lngIndex, 5) = oProd.OSAPs(lngIndex).TPNAME
        XE.Value(lngIndex, 3) = oProd.OSAPs(lngIndex).APPQty
        XE.Value(lngIndex, 4) = oProd.OSAPs(lngIndex).APPReturned
        XE.Value(lngIndex, 6) = oProd.OSAPs(lngIndex).TRID
        XE.Value(lngIndex, 7) = oProd.OSAPs(lngIndex).DateForSort
    Next
    XE.QuickSort 1, oProd.OSAPs.Count, 7, XORDER_DESCEND, XTYPE_STRING
    APPGRID.Array = XE
    APPGRID.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.LoadAPs"
End Sub
Private Sub LoadMMs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XF = New XArrayDB
    XF.Clear
    XF.ReDim 1, oProd.MMs.Count, 1, 7
    For lngIndex = 1 To oProd.MMs.Count
        XF.Value(lngIndex, 1) = oProd.MMs(lngIndex).DOCCode
        XF.Value(lngIndex, 2) = oProd.MMs(lngIndex).DocDateF
        XF.Value(lngIndex, 3) = oProd.MMs(lngIndex).Qty
        XF.Value(lngIndex, 4) = oProd.MMs(lngIndex).typ
        XF.Value(lngIndex, 5) = oProd.MMs(lngIndex).PID
        XF.Value(lngIndex, 6) = oProd.MMs(lngIndex).TRID
        XF.Value(lngIndex, 7) = oProd.MMs(lngIndex).DateForSort
    Next
    XF.QuickSort 1, oProd.MMs.Count, 7, XORDER_DESCEND, XTYPE_STRING
    MMGRID.Array = XF
    MMGRID.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNBPrev.LoadMMs"
End Sub



