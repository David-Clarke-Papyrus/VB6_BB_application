VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmProductSinglePreview 
   BackColor       =   &H00D3D3CB&
   Caption         =   "General stock"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   14445
   ControlBox      =   0   'False
   Icon            =   "frmProductSinglesPreview.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleMode       =   0  'User
   ScaleWidth      =   19055.11
   Begin VB.ListBox lbPSECs 
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      ItemData        =   "frmProductSinglesPreview.frx":030A
      Left            =   5760
      List            =   "frmProductSinglesPreview.frx":030C
      TabIndex        =   92
      Top             =   1080
      Width           =   1740
   End
   Begin VB.TextBox txtProductType 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   90
      Top             =   435
      Width           =   1725
   End
   Begin VB.CommandButton cmdForward 
      Height          =   390
      Left            =   4095
      Picture         =   "frmProductSinglesPreview.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   5970
      Width           =   435
   End
   Begin VB.CommandButton cmdBack 
      Height          =   390
      Left            =   3585
      Picture         =   "frmProductSinglesPreview.frx":0698
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   5970
      Width           =   435
   End
   Begin TrueOleDBGrid60.TDBGrid StGrid 
      Height          =   1485
      Left            =   3600
      OleObjectBlob   =   "frmProductSinglesPreview.frx":0A22
      TabIndex        =   50
      Top             =   6660
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.TextBox txtLastOrdered 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   7035
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtLastOrderedQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   7035
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtLastOrderedPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6990
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   7035
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txtLastReceivedPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6990
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   7350
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txtLastReceivedQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   7350
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtLastReceived 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   7350
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox txtLastSoldPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6990
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   7665
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.TextBox txtLastSoldQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   7665
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtLastSoldDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   7665
      Visible         =   0   'False
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
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8010
      Visible         =   0   'False
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
      Picture         =   "frmProductSinglesPreview.frx":4C65
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
      Picture         =   "frmProductSinglesPreview.frx":4D10
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
      Left            =   180
      TabIndex        =   8
      Top             =   6000
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
      Height          =   4020
      Left            =   60
      TabIndex        =   4
      Top             =   1890
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   7091
      _Version        =   393216
      Style           =   1
      Tabs            =   2
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
      TabPicture(0)   =   "frmProductSinglesPreview.frx":501A
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
      Tab(0).Control(9)=   "Label2(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2(2)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label12"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label22"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "MMGRID"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "APPGRID"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtOnHand"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtReserved"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtTotalSold"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtTotalOSAP"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdALLMM"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "DTMMSince"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtReturnable"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdRecon"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtSSP"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtCost"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtSP"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtRRP"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtCategorization(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtCategorization(1)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtCategorization(2)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtCategorization(3)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtCategorization(4)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtCategorization(5)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtCategorization(6)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtCategorization(7)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtDateAdded"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtDateLastModified"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProductSinglesPreview.frx":5036
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label26"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblObsolete"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblSupplier"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label23"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label33"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label34"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtVAT"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtCategory"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtEdition"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtPublisher"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtSS"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtDefaultDeliveryDays"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).ControlCount=   15
      Begin VB.TextBox txtDateLastModified 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   3075
         Width           =   1890
      End
      Begin VB.TextBox txtDateAdded 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   2655
         Width           =   1890
      End
      Begin VB.TextBox txtCategorization 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   7
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   3180
         Width           =   2370
      End
      Begin VB.TextBox txtCategorization 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   6
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   2793
         Width           =   2370
      End
      Begin VB.TextBox txtCategorization 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   5
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   2410
         Width           =   2370
      End
      Begin VB.TextBox txtCategorization 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   4
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   81
         Top             =   2027
         Width           =   2370
      End
      Begin VB.TextBox txtCategorization 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   3
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   1644
         Width           =   2370
      End
      Begin VB.TextBox txtCategorization 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   2
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   1261
         Width           =   2370
      End
      Begin VB.TextBox txtCategorization 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   1
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   878
         Width           =   2370
      End
      Begin VB.TextBox txtCategorization 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   0
         Left            =   4485
         Locked          =   -1  'True
         TabIndex        =   77
         Top             =   495
         Width           =   2370
      End
      Begin VB.TextBox txtRRP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2220
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.TextBox txtSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1530
         Width           =   1200
      End
      Begin VB.TextBox txtCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2160
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
         TabIndex        =   42
         Top             =   2235
         Visible         =   0   'False
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
         TabIndex        =   41
         Top             =   3510
         Width           =   720
      End
      Begin VB.TextBox txtReturnable 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   3270
         Visible         =   0   'False
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker DTMMSince 
         Height          =   360
         Left            =   9675
         TabIndex        =   36
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
         Format          =   16515073
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
         TabIndex        =   35
         Top             =   3510
         Width           =   1695
      End
      Begin VB.TextBox txtDefaultDeliveryDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73095
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3360
         Visible         =   0   'False
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
         TabIndex        =   29
         Top             =   2515
         Visible         =   0   'False
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
         TabIndex        =   26
         Top             =   1680
         Width           =   3705
      End
      Begin VB.TextBox txtPublisher 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1255
         Width           =   3990
      End
      Begin VB.TextBox txtEdition 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1675
         Visible         =   0   'False
         Width           =   3990
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
      Begin TrueOleDBGrid60.TDBGrid APPGRID 
         Height          =   1245
         Left            =   7110
         OleObjectBlob   =   "frmProductSinglesPreview.frx":5052
         TabIndex        =   27
         Top             =   435
         Width           =   4065
      End
      Begin TrueOleDBGrid60.TDBGrid MMGRID 
         Height          =   1215
         Left            =   7155
         OleObjectBlob   =   "frmProductSinglesPreview.frx":94EE
         TabIndex        =   28
         Top             =   2250
         Width           =   4035
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Last modified"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   15
         TabIndex        =   89
         Top             =   3105
         Width           =   930
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Added"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   105
         TabIndex        =   88
         Top             =   2715
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock group"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Index           =   0
         Left            =   2910
         TabIndex        =   85
         Top             =   495
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock group"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   7
         Left            =   2910
         TabIndex        =   76
         Top             =   3195
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock group"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   6
         Left            =   2910
         TabIndex        =   75
         Top             =   2799
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock group"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   5
         Left            =   2910
         TabIndex        =   74
         Top             =   2405
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock group"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   4
         Left            =   2910
         TabIndex        =   73
         Top             =   2011
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock group"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   3
         Left            =   2910
         TabIndex        =   72
         Top             =   1605
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock group"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   2
         Left            =   2910
         TabIndex        =   71
         Top             =   1223
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Stock group"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   1
         Left            =   2910
         TabIndex        =   70
         Top             =   829
         Width           =   1485
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R.R.P."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2400
         TabIndex        =   49
         Top             =   1995
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sell.P."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   240
         TabIndex        =   48
         Top             =   1290
         Width           =   540
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cost (avg.)"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   150
         TabIndex        =   47
         Top             =   1920
         Width           =   1110
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Special"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1665
         TabIndex        =   46
         Top             =   2010
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Returnable"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   225
         TabIndex        =   39
         Top             =   3315
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   37
         Top             =   3645
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Lead time"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74775
         TabIndex        =   34
         Top             =   3390
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Last supplied by"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   -74775
         TabIndex        =   32
         Top             =   2992
         Width           =   1590
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Order by seesafe"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74775
         TabIndex        =   30
         Top             =   2555
         Visible         =   0   'False
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   1711
         Visible         =   0   'False
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
      Height          =   450
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1125
      Width           =   5265
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   660
      Left            =   75
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   5250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5970
      TabIndex        =   93
      Top             =   855
      Width           =   1080
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Product type"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5970
      TabIndex        =   91
      Top             =   195
      Width           =   1080
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1920
      Left            =   7980
      Top             =   60
      Width           =   3465
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last ord'd"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   4425
      TabIndex        =   67
      Top             =   7065
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last rec'd"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   4320
      TabIndex        =   66
      Top             =   7380
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   5475
      TabIndex        =   65
      Top             =   6840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   6270
      TabIndex        =   64
      Top             =   6840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   7005
      TabIndex        =   63
      Top             =   6840
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last sold"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   4515
      TabIndex        =   62
      Top             =   7680
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "(RRP)"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   10575
      TabIndex        =   61
      Top             =   6930
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "(Cost)"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   10575
      TabIndex        =   60
      Top             =   7215
      Visible         =   0   'False
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
      TabIndex        =   31
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
      Left            =   45
      TabIndex        =   3
      Top             =   930
      Width           =   645
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   30
      Width           =   945
   End
End
Attribute VB_Name = "frmProductSinglePreview"
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
Dim tlProductCategorizations As New z_TextList
Dim tlCollection As Collection

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
    ErrorIn "frmProductSinglePreview.component(pProduct,pstrTime)", Array(pProduct, pstrTime)
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
    ErrorIn "frmProductSinglePreview.cmdExpand_Click", , EA_NORERAISE
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
    ErrorIn "frmProductSinglePreview.cmdBack_Click", , EA_NORERAISE
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
    ErrorIn "frmProductSinglePreview.cmdForward_Click", , EA_NORERAISE
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
    ErrorIn "frmProductSinglePreview.cmdRecon_Click", , EA_NORERAISE
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
    ErrorIn "frmProductSinglePreview.cmdSales_Click", , EA_NORERAISE
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
    ErrorIn "frmProductSinglePreview.cmdSub_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdSubs_Click()
'  '  Set XSF = New XArrayDB
'  '  Set XSB = New XArrayDB
'    oProd.GetSubstitutes XSF, XSB
'  '  Set GSF.Array = XSF
'   ' Set GSB.Array = XSB
'    GSF.ReBind
'    GSB.ReBind
'End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuAdjust.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.SetMenu"
End Sub

Private Sub cboCatHead_Click()
    On Error GoTo errHandler
'    oProd.setCatalogueheadingID tlCatHead.Key(cboCatHead)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.cboCatHead_Click", , EA_NORERAISE
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
    ErrorIn "frmProductSinglePreview.cmdDelete_Click", , EA_NORERAISE
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
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.APPGRID_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_AxisLabelSelected(axisID As Integer, AxisIndex As Integer, labelSetIndex As Integer, LabelIndex As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.chart1_AxisLabelSelected(axisID,AxisIndex,labelSetIndex," & _
        "LabelIndex,MouseFlags,Cancel)", Array(axisID, AxisIndex, labelSetIndex, LabelIndex, MouseFlags, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_AxisTitleSelected(axisID As Integer, AxisIndex As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.chart1_AxisTitleSelected(axisID,AxisIndex,MouseFlags,Cancel)", _
         Array(axisID, AxisIndex, MouseFlags, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_ChartSelected(MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.chart1_ChartSelected(MouseFlags,Cancel)", Array(MouseFlags, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_FootnoteSelected(MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.chart1_FootnoteSelected(MouseFlags,Cancel)", Array(MouseFlags, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_LegendSelected(MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.chart1_LegendSelected(MouseFlags,Cancel)", Array(MouseFlags, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_PlotSelected(MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.chart1_PlotSelected(MouseFlags,Cancel)", Array(MouseFlags, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
'MsgBox "Series " & Series
'MsgBox "Datapoint " & DataPoint
    Select Case Series
    Case 1
        MsgBox "Sales qty = " & oProd.CurrentSales.FindByWeek(DataPoint).qty & vbCrLf & "Sales value = " & oProd.CurrentSales.FindByWeek(DataPoint).ValuF
    Case 2
   '     MsgBox "Sales qty = " & oProd.PreviousSales.FindByWeek(DataPoint).Qty & vbCrLf & "Sales value = " & oProd.PreviousSales.FindByWeek(DataPoint).ValuF
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.chart1_PointSelected(Series,DataPoint,MouseFlags,Cancel)", _
         Array(Series, DataPoint, MouseFlags, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub chart1_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
    On Error GoTo errHandler
'Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.chart1_SeriesSelected(Series,MouseFlags,Cancel)", Array(Series, _
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
    ErrorIn "frmProductSinglePreview.cmdALLMM_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdclose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmProductSingles
    Set frm = New frmProductSingles
    frm.component oProd, Me
    frm.Show
    Exit Sub
Errh:
    MsgBox Error

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.cmdEdit_Click", , EA_NORERAISE
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
    ErrorIn "frmProductSinglePreview.cmdsearchisbn_Click", , EA_NORERAISE
    HandleError
End Sub





Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
            Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
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
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.MMGRID_DblClick", , EA_NORERAISE
    HandleError
End Sub

'Public Sub mnuAdjust()
'Dim frm As New frmStockAdjust
'    frm.Component oProd
'    frm.Show vbModal
'    If frm.Cancelled = False Then
'
'        Me.txtLastCounted = Now
'        Me.txtLastCountedQty = frm.Counted
'        Me.txtOnHand = frm.Count
'    End If
'    Unload frm
'    Unload Me
'End Sub
'


Private Sub StGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.StGrid_BeforeColUpdate(ColIndex,OldValue,Cancel)", _
         Array(ColIndex, OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        top = 10
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
    ErrorIn "frmProductSinglePreview.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Public Sub RefreshForm()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.RefreshForm"
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
Dim i As Integer

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
  '  Me.txtAgedDate = ""
    Me.txtCost = oProd.CostF
    Me.txtDateAdded = oProd.DateRecordAddedF
    Me.txtDateLastModified = oProd.DateLastModifiedF
 '   Me.txtLastCounted = oProd.dateLastCountedF
'   Me.txtLastCountedPrice = oProd.PriceLastCountedF
 '   Me.txtLastCountedQty = oProd.QtyLastCountedF
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
    lblStatus = oProd.statusF
    DTMMSince.Value = IIf(oProd.DateLastCounted < DTMMSince.MinDate, DTMMSince.MinDate, oProd.DateLastCounted)
    LoadMovements

    LoadStock
    LoadProductSections

    Set tlCollection = New Collection
    For i = 1 To 8
        tlCollection.Add New z_TextList
    Next
    For i = 1 To 8
        Me.txtCategorization(i - 1).Visible = False
        Me.Label2(i - 1).Visible = False
    Next
    
    Set tlProductCategorizations = New z_TextList
    tlProductCategorizations.Load ltProductCategorizations
    For i = 0 To tlProductCategorizations.Count - 1
      ' LoadCombo txtCategorization(i), GetTextList(i)
        Me.txtCategorization(i).Visible = True
        Me.Label2(i).Visible = True
        Me.Label2(i).Caption = tlProductCategorizations.ItemByOrdinalIndex(i + 1)
        If tlCollection(i + 1).ItemByOrdinalIndex(1) > "" Then txtCategorization(i).Text = tlCollection(i + 1).ItemByOrdinalIndex(1)
        If Not oProd.ProductCategories.ItemByCatID(CStr(tlProductCategorizations.KeyByOrdinalIndex(i + 1))) Is Nothing Then
            txtCategorization(i).Text = oProd.ProductCategories.ItemByCatID(CStr(tlProductCategorizations.KeyByOrdinalIndex(i + 1))).Description
        End If
    Next

    LoadPictureFromDB oProd.PID


    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.LoadControls"
End Sub
Private Sub LoadPictureFromDB(PID As String)
    On Error GoTo errHandler
Dim bytTemp() As Byte
    bytTemp = ImageFromDB(PID)
    If UBound(bytTemp) > 0 Then
        Image1.Stretch = True
        Image1.Width = 3000
        Image1.Height = 1500
        Set Image1 = ArrayToPictureB(bytTemp(), 0, UBound(bytTemp) + 1)
    Else
        Set Image1.Picture = LoadPicture
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.LoadPictureFromDB(PID)", PID
End Sub

Private Function GetTextList(i As Integer) As z_TextList
    On Error GoTo errHandler
        tlCollection(i + 1).Load ltProductCategorizationValues, CStr(tlProductCategorizations.KeyByOrdinalIndex(i + 1)), "<n/a>"
        Set GetTextList = tlCollection(i + 1)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.GetTextList(i)", i
End Function

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
    ErrorIn "frmProductSinglePreview.LoadProductSections"
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
        XB.Value(lngIndex, 1) = oProd.Stores(lngIndex).Storename
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
    ErrorIn "frmProductSinglePreview.LoadStock"
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
    
    
    APPGRID.Splits(0).Caption = "Appros issued"
    MMGRID.Splits(0).Caption = "Movements since last count (" & oProd.dateLastCountedF & ":" & oProd.QtyLastCountedF & ")"
    
    txtTotalOSAP = "Expecting return of " & lngQtyApps & " copies."
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSinglePreview.LoadMovements"
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
        XE.Value(lngIndex, 1) = oProd.OSAPs(lngIndex).DocCode
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
    ErrorIn "frmProductSinglePreview.LoadAPs"
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
        XF.Value(lngIndex, 1) = oProd.MMs(lngIndex).DocCode
        XF.Value(lngIndex, 2) = oProd.MMs(lngIndex).DocDateF
        XF.Value(lngIndex, 3) = oProd.MMs(lngIndex).qty
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
    ErrorIn "frmProductSinglePreview.LoadMMs"
End Sub



