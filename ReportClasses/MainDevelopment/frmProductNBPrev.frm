VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
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
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleMode       =   0  'User
   ScaleWidth      =   15592.34
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
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   78
      Top             =   1320
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
      Picture         =   "frmProductNBPrev.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5505
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
      Picture         =   "frmProductNBPrev.frx":03B5
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5505
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Find By ISBN"
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   240
      TabIndex        =   8
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
      Top             =   1455
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   7064
      _Version        =   393216
      Style           =   1
      Tab             =   2
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
      TabPicture(0)   =   "frmProductNBPrev.frx":06BF
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtRRP"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtSP"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCost"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtSSP"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdRecon"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtReturnable"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DTMMSince"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdALLMM"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtTotalOSAP"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtTotalOSCO"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtTotalOSPO"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtTotalSold"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtReserved"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtOnHand"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "POGrid"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "COGrid"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "APPGRID"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "MMGRID"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label16"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label18"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label19"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label29"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label40"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblStatus"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label21"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label14"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label13"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProductNBPrev.frx":06DB
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
      TabPicture(2)   =   "frmProductNBPrev.frx":06F7
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label39"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label38"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label37"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label36"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label15"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label24"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label22"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label12"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame2"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtLastSoldDate"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtLastSoldQty"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtLastSoldPrice"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtLastCountedPrice"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtLastCountedQty"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtLastCounted"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtLastReceived"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "txtLastReceivedQty"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtLastReceivedPrice"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtLastOrderedPrice"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtLastOrderedQty"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtLastOrdered"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txtDateLastModified"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtDateAdded"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).ControlCount=   24
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
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   1635
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
         Left            =   -73275
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   1635
         Width           =   1200
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
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   2235
         Width           =   1170
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
         Left            =   -73275
         Locked          =   -1  'True
         TabIndex        =   80
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
         Left            =   -67770
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   3510
         Width           =   645
      End
      Begin VB.ListBox lbPSECs 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   810
         Left            =   -66870
         TabIndex        =   76
         Top             =   1140
         Width           =   2715
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
         Left            =   8955
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   3060
         Width           =   2160
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
         Left            =   8955
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   3480
         Width           =   2160
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   1845
         Width           =   1380
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
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1845
         Width           =   555
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1845
         Width           =   960
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   2265
         Width           =   960
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
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   2265
         Width           =   555
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
         Left            =   1905
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   2265
         Width           =   1380
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   2685
         Width           =   1380
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
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   2685
         Width           =   555
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   2685
         Width           =   960
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
         Left            =   3945
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   3105
         Width           =   960
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
         Left            =   3345
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3105
         Width           =   555
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3105
         Width           =   1380
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
         Left            =   435
         TabIndex        =   47
         Top             =   570
         Width           =   5220
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
            TabIndex        =   52
            Top             =   330
            Width           =   1170
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   330
            Width           =   825
         End
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
            TabIndex        =   48
            Top             =   330
            Width           =   825
         End
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
         Left            =   -73755
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   3270
         Visible         =   0   'False
         Width           =   1380
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
         Left            =   -66870
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   630
         Width           =   2370
      End
      Begin MSComCtl2.DTPicker DTMMSince 
         Height          =   360
         Left            =   -65325
         TabIndex        =   41
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
         Format          =   49938433
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
         Left            =   -67080
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3510
         Width           =   1695
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
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3360
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
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2515
         Width           =   1395
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
         Left            =   -67740
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1680
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
         Left            =   -71925
         Locked          =   -1  'True
         TabIndex        =   28
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
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1255
         Width           =   3990
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
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   1675
         Width           =   3990
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
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1650
         Width           =   3705
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
         Left            =   -72870
         TabIndex        =   15
         Top             =   825
         Width           =   750
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
         Left            =   -73830
         Locked          =   -1  'True
         TabIndex        =   14
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
         Left            =   -74790
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   825
         Width           =   750
      End
      Begin VB.TextBox txtCategory 
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
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2095
         Width           =   1395
      End
      Begin TrueOleDBGrid60.TDBGrid POGrid 
         Height          =   1230
         Left            =   -71925
         OleObjectBlob   =   "frmProductNBPrev.frx":0713
         TabIndex        =   22
         Top             =   420
         Width           =   3960
      End
      Begin TrueOleDBGrid60.TDBGrid COGrid 
         Height          =   1245
         Left            =   -71910
         OleObjectBlob   =   "frmProductNBPrev.frx":5066
         TabIndex        =   29
         Top             =   2235
         Width           =   3945
      End
      Begin TrueOleDBGrid60.TDBGrid APPGRID 
         Height          =   1245
         Left            =   -67890
         OleObjectBlob   =   "frmProductNBPrev.frx":9501
         TabIndex        =   31
         Top             =   435
         Width           =   4065
      End
      Begin TrueOleDBGrid60.TDBGrid MMGRID 
         Height          =   1215
         Left            =   -67845
         OleObjectBlob   =   "frmProductNBPrev.frx":D99D
         TabIndex        =   33
         Top             =   2250
         Width           =   4035
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
         Left            =   -74730
         TabIndex        =   87
         Top             =   1410
         Width           =   540
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
         Left            =   -73260
         TabIndex        =   86
         Top             =   1410
         Width           =   540
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
         Left            =   -74850
         TabIndex        =   85
         Top             =   1995
         Width           =   1110
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
         Left            =   -73335
         TabIndex        =   84
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label Label2 
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
         Left            =   -68130
         TabIndex        =   77
         Top             =   1140
         Width           =   1080
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
         Left            =   7560
         TabIndex        =   75
         Top             =   3135
         Width           =   1290
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
         Left            =   7470
         TabIndex        =   74
         Top             =   3525
         Width           =   1395
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
         Left            =   375
         TabIndex        =   73
         Top             =   1875
         Width           =   1395
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
         Left            =   375
         TabIndex        =   72
         Top             =   2295
         Width           =   1395
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
         Left            =   375
         TabIndex        =   71
         Top             =   2715
         Width           =   1395
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
         Left            =   2145
         TabIndex        =   70
         Top             =   1560
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
         Left            =   3195
         TabIndex        =   69
         Top             =   1560
         Width           =   840
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
         Left            =   4050
         TabIndex        =   68
         Top             =   1560
         Width           =   840
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
         Left            =   375
         TabIndex        =   67
         Top             =   3135
         Width           =   1395
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
         Left            =   -74790
         TabIndex        =   46
         Top             =   3315
         Visible         =   0   'False
         Width           =   975
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
         Left            =   -68130
         TabIndex        =   44
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   -73770
         TabIndex        =   42
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
         Left            =   -74775
         TabIndex        =   39
         Top             =   3390
         Width           =   1590
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Last supplied by"
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
         Left            =   -74775
         TabIndex        =   37
         Top             =   2992
         Width           =   1590
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
         Left            =   -74775
         TabIndex        =   35
         Top             =   2555
         Width           =   1590
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "Manufacturer"
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
         Left            =   -74385
         TabIndex        =   27
         Top             =   1305
         Width           =   1200
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         Caption         =   "Version"
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
         Left            =   -73830
         TabIndex        =   26
         Top             =   1711
         Width           =   645
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -73080
         TabIndex        =   19
         Top             =   2935
         Width           =   3135
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
         Left            =   -72960
         TabIndex        =   18
         Top             =   480
         Width           =   870
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
         Left            =   -73950
         TabIndex        =   17
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
         Left            =   -74775
         TabIndex        =   16
         Top             =   480
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
         Left            =   -74265
         TabIndex        =   7
         Top             =   855
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
         Left            =   -74265
         TabIndex        =   6
         Top             =   2118
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
      Left            =   1155
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   615
      Width           =   5250
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
      Height          =   375
      Left            =   1155
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   5355
   End
   Begin TrueOleDBGrid60.TDBGrid StGrid 
      Height          =   1200
      Left            =   6645
      OleObjectBlob   =   "frmProductNBPrev.frx":11AA0
      TabIndex        =   32
      Top             =   90
      Width           =   4815
   End
   Begin VB.Label lblNonStock 
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
      Left            =   1140
      TabIndex        =   36
      Top             =   1050
      Width           =   2460
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
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
      Left            =   405
      TabIndex        =   3
      Top             =   645
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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

'    Set tlCatHead = Nothing
'    Set tlCatHead = New z_TextList
'    tlCatHead.Load ltCatalogueHeadings
strTime = strTime & "End frmProductPrev:component:" & Now() & vbCrLf
End Sub

Private Sub cmdExpand_Click()
Dim frm As frmSalesCH
    Set frm = New frmSalesCH
    frm.Component oProd
    frm.Show vbModal
End Sub

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

Private Sub Form_Deactivate()
    UnsetMenu
End Sub
Private Sub Form_Activate()
    SetMenu
End Sub
Private Sub SetMenu()
    Forms(0).mnuAdjust.Enabled = True
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
Dim frmMM As frmMovements
    oProd.ReloadMovements Me.DTMMSince
    Set frmMM = New frmMovements
    frmMM.Component oProd
    frmMM.Show
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
Dim frm As frmProductNB
    Set frm = New frmProductNB
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
       
    Me.Caption = "Stock code: " & .code
    txtSubtitle = .SubTitle
    txtTitle = .Title
    txtPublisher = .Publisher
        
    End With
    LoadControls
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




Private Sub MMGRID_DblClick()
Dim strType As String
Dim frm As Form
Dim i As Integer


    If IsNull(MMGRID.Bookmark) Then Exit Sub
    Screen.MousePointer = vbHourglass

    strType = XF(MMGRID.Bookmark, 4)
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
'        If Not ISOpenForm(frmInvoicePreview, i) Then
            Set frm = New frmInvoicePreview
            frm.Component XF(MMGRID.Bookmark, 6)
            frm.Show
'        Else
'            Set frm = Forms(i)
'            frm.ZOrder 0
'        End If
    Case "CS"
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
'    Case "RET"
'        Set frm = New frmReturn3
'        Set oR = new d_R
'
'        frm.Component XF(MMGRID.Bookmark, 6), ""
'        frm.Show
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


Private Sub StGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Cancel = True
End Sub

Private Sub Form_Load()
    left = 10
    top = 10
    Width = 11700
    Height = 6800
strTime = strTime & "Start frmProductPrev:Load:" & Now() & vbCrLf
    LoadControls
strTime = strTime & "End frmProductPrev:Load:" & Now() & vbCrLf
    Me.SSTab1.Tab = 0
End Sub
Public Sub RefreshForm()
    LoadControls
End Sub
Private Sub LoadControls()
    flgLoading = True
    Me.Caption = "Stock code: " & oProd.code & "      EAN: " & oProd.Ean
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
    Me.lblNonStock.Visible = oProd.IsNONStock
    Me.lblObsolete = IIf(oProd.Obsolete, "obsolete", "")
    txtSS = IIf(oProd.Seesafe = 1, "Yes", "")
    Me.lblSupplier.Caption = oProd.lastsuppliername
    lblStatus = oProd.statusF
    DTMMSince.Value = IIf(oProd.DateLastCounted < DTMMSince.MinDate, DTMMSince.MinDate, oProd.DateLastCounted)
    LoadMovements

    LoadStock
    LoadProductSections

    flgLoading = False
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

Private Sub LoadStock()
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
        XB.Value(lngIndex, 5) = oProd.Stores(lngIndex).QtyOnOrder
        XB.Value(lngIndex, 6) = oProd.Stores(lngIndex).QtyCopiesOnHand
    Next
    XB.QuickSort 1, oProd.Stores.Count, 1, XORDER_ASCEND, XTYPE_STRING
    StGrid.Array = XB
    StGrid.ReBind

End Sub

Private Sub LoadMovements()
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
        XE.Value(lngIndex, 7) = oProd.OSAPs(lngIndex).DateForSort
    Next
    XE.QuickSort 1, oProd.OSAPs.Count, 7, XORDER_DESCEND, XTYPE_STRING
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
        XF.Value(lngIndex, 7) = oProd.MMs(lngIndex).DateForSort
    Next
    XF.QuickSort 1, oProd.MMs.Count, 7, XORDER_DESCEND, XTYPE_STRING
    MMGRID.Array = XF
    MMGRID.ReBind

End Sub



