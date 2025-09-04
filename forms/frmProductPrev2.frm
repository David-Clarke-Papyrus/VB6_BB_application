VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Begin VB.Form frmProductPrev 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Stock"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   11820
   ControlBox      =   0   'False
   Icon            =   "frmProductPrev2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   11820
   Begin VB.CommandButton cmdWsSales 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Wordstock sales"
      Height          =   525
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   140
      Top             =   1095
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TrueOleDBGrid60.TDBGrid StGrid 
      Height          =   1485
      Left            =   9885
      OleObjectBlob   =   "frmProductPrev2.frx":030A
      TabIndex        =   121
      Top             =   60
      Width           =   1650
   End
   Begin VB.TextBox txtLastSoldDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   131
      Top             =   840
      Width           =   960
   End
   Begin VB.TextBox txtLastSoldQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   130
      Top             =   840
      Width           =   555
   End
   Begin VB.TextBox txtLastSoldPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8070
      Locked          =   -1  'True
      TabIndex        =   129
      Top             =   840
      Width           =   885
   End
   Begin VB.TextBox txtLastReceived 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   128
      Top             =   525
      Width           =   960
   End
   Begin VB.TextBox txtLastReceivedQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   127
      Top             =   525
      Width           =   555
   End
   Begin VB.TextBox txtLastReceivedPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8070
      Locked          =   -1  'True
      TabIndex        =   126
      Top             =   525
      Width           =   885
   End
   Begin VB.TextBox txtLastOrderedPrice 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8070
      Locked          =   -1  'True
      TabIndex        =   125
      Top             =   210
      Width           =   885
   End
   Begin VB.TextBox txtLastOrderedQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   124
      Top             =   210
      Width           =   555
   End
   Begin VB.TextBox txtLastOrdered 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6510
      Locked          =   -1  'True
      TabIndex        =   123
      Top             =   210
      Width           =   960
   End
   Begin VB.CommandButton cmdDropStock 
      BackColor       =   &H00C4BCA4&
      Height          =   315
      Left            =   11520
      Picture         =   "frmProductPrev2.frx":454D
      Style           =   1  'Graphical
      TabIndex        =   122
      Top             =   1230
      Width           =   270
   End
   Begin VB.CommandButton cmdWash 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Wash with Nielsen"
      Height          =   360
      Left            =   6555
      Style           =   1  'Graphical
      TabIndex        =   115
      Top             =   1260
      Width           =   1485
   End
   Begin VB.CommandButton cmdBack 
      Height          =   390
      Left            =   3675
      Picture         =   "frmProductPrev2.frx":48D7
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   6075
      Width           =   435
   End
   Begin VB.CommandButton cmdForward 
      Height          =   390
      Left            =   4185
      Picture         =   "frmProductPrev2.frx":4C61
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   6075
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Sales"
      Enabled         =   0   'False
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
      TabIndex        =   110
      Top             =   7440
      Width           =   1200
   End
   Begin VB.CommandButton cmdSales 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8055
      Style           =   1  'Graphical
      TabIndex        =   100
      Top             =   1260
      Width           =   840
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
      Height          =   615
      Left            =   10605
      Picture         =   "frmProductPrev2.frx":4FEB
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6015
      Width           =   1000
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
      Height          =   615
      Left            =   9585
      Picture         =   "frmProductPrev2.frx":5375
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6015
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Find By ISBN"
      ForeColor       =   &H00800000&
      Height          =   870
      Left            =   240
      TabIndex        =   12
      Top             =   5985
      Width           =   3375
      Begin VB.CommandButton cmdsearchisbn 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2220
         Picture         =   "frmProductPrev2.frx":56FF
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   195
         Width           =   1005
      End
      Begin VB.TextBox txtisbnsearch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   150
         TabIndex        =   13
         Top             =   270
         Width           =   1995
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4515
      Left            =   225
      TabIndex        =   6
      Top             =   1410
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   7964
      _Version        =   393216
      Style           =   1
      Tabs            =   6
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
      TabPicture(0)   =   "frmProductPrev2.frx":5A89
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
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
      Tab(0).Control(12)=   "Label3"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "MMGRID"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "APPGRID"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "COGrid"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "POGrid"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtOnHand"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtReserved"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtRRP"
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
      Tab(0).Control(32)=   "txtEUPrice"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdDropPO"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdDropCO"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdDropMovements"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmdDropAppros"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtSP"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmdStatusChange"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtWeight"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdExplainCost"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "chkCore"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdDefaultColWidths"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProductPrev2.frx":5AA5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label26"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblDeal"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblSupplier"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblPublicationPlace"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblPublicationDate"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblEdition"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblPublisher"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label23"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label27"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label33"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label34"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label35"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label11"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtBinding"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtVAT"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtPubPlace"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtPubDate"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtEdition"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtPublisher"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtSS"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtDefaultDeliveryDays"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "txtProductType"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Frame3"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "txtLoyaltyRate"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "lbPSECs"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "&3. Notes"
      TabPicture(2)   =   "frmProductPrev2.frx":5AC1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtNotes"
      Tab(2).Control(1)=   "txtCatalogues"
      Tab(2).Control(2)=   "txtCategoryHeading"
      Tab(2).Control(3)=   "txtFlagText"
      Tab(2).Control(4)=   "txtComment"
      Tab(2).Control(5)=   "txtDescription"
      Tab(2).Control(6)=   "Label7"
      Tab(2).Control(7)=   "Label25"
      Tab(2).Control(8)=   "Label17"
      Tab(2).Control(9)=   "Label2"
      Tab(2).Control(10)=   "Label30"
      Tab(2).Control(11)=   "Label28"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "&4. Statistics"
      TabPicture(3)   =   "frmProductPrev2.frx":5ADD
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtLastCountedPrice"
      Tab(3).Control(1)=   "txtLastCountedQty"
      Tab(3).Control(2)=   "txtLastCounted"
      Tab(3).Control(3)=   "txtDateLastModified"
      Tab(3).Control(4)=   "txtDateAdded"
      Tab(3).Control(5)=   "Frame2"
      Tab(3).Control(6)=   "Label45"
      Tab(3).Control(7)=   "Shape1"
      Tab(3).Control(8)=   "Label38"
      Tab(3).Control(9)=   "Label37"
      Tab(3).Control(10)=   "Label36"
      Tab(3).Control(11)=   "Label1"
      Tab(3).Control(12)=   "Label22"
      Tab(3).Control(13)=   "Label12"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "&5. Substitute products"
      TabPicture(4)   =   "frmProductPrev2.frx":5AF9
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdSub"
      Tab(4).Control(1)=   "cmdSubs"
      Tab(4).Control(2)=   "GSB"
      Tab(4).Control(3)=   "GSF"
      Tab(4).Control(4)=   "Label43"
      Tab(4).Control(5)=   "Label44"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "&6. Copies"
      TabPicture(5)   =   "frmProductPrev2.frx":5B15
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grid1"
      Tab(5).ControlCount=   1
      Begin VB.CommandButton cmdDefaultColWidths 
         BackColor       =   &H00C4BCA4&
         Height          =   315
         Left            =   7065
         Picture         =   "frmProductPrev2.frx":5B31
         Style           =   1  'Graphical
         TabIndex        =   147
         ToolTipText     =   "Set column widths to default"
         Top             =   2145
         Width           =   330
      End
      Begin VB.CheckBox chkCore 
         Alignment       =   1  'Right Justify
         Caption         =   "Core"
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   2115
         TabIndex        =   146
         Top             =   3375
         Width           =   780
      End
      Begin VB.CommandButton cmdExplainCost 
         Caption         =   "?"
         Height          =   285
         Left            =   1365
         TabIndex        =   145
         Top             =   1920
         Width           =   195
      End
      Begin VB.TextBox txtWeight 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   144
         Top             =   3735
         Width           =   855
      End
      Begin VB.CommandButton cmdStatusChange 
         BackColor       =   &H00C4BCA4&
         Caption         =   "Status"
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
         Left            =   2145
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   4095
         Width           =   645
      End
      Begin VB.TextBox txtNotes 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   825
         Left            =   -67170
         MultiLine       =   -1  'True
         TabIndex        =   141
         Top             =   2505
         Width           =   3240
      End
      Begin VB.TextBox txtSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdDropAppros 
         BackColor       =   &H00C4BCA4&
         Height          =   315
         Left            =   11130
         Picture         =   "frmProductPrev2.frx":5EBB
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   1680
         Width           =   225
      End
      Begin VB.CommandButton cmdDropMovements 
         BackColor       =   &H00C4BCA4&
         Height          =   315
         Left            =   11130
         Picture         =   "frmProductPrev2.frx":6245
         Style           =   1  'Graphical
         TabIndex        =   119
         Top             =   3645
         Width           =   225
      End
      Begin VB.CommandButton cmdDropCO 
         BackColor       =   &H00C4BCA4&
         Height          =   315
         Left            =   7095
         Picture         =   "frmProductPrev2.frx":65CF
         Style           =   1  'Graphical
         TabIndex        =   118
         Top             =   3660
         Width           =   225
      End
      Begin VB.CommandButton cmdDropPO 
         BackColor       =   &H00C4BCA4&
         Height          =   315
         Left            =   7110
         Picture         =   "frmProductPrev2.frx":6959
         Style           =   1  'Graphical
         TabIndex        =   117
         Top             =   1680
         Width           =   225
      End
      Begin VB.TextBox txtEUPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   3735
         Width           =   855
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
         Left            =   7410
         Style           =   1  'Graphical
         TabIndex        =   111
         Top             =   4020
         Width           =   645
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
         TabIndex        =   48
         Top             =   450
         Width           =   1920
      End
      Begin VB.TextBox txtSSP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   1905
         Width           =   1200
      End
      Begin VB.TextBox txtCatalogues 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   1395
         Left            =   -67170
         MultiLine       =   -1  'True
         TabIndex        =   98
         Top             =   690
         Width           =   3240
      End
      Begin VB.ListBox lbPSECs 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   615
         Left            =   -73620
         TabIndex        =   96
         Top             =   885
         Width           =   2715
      End
      Begin VB.TextBox txtLoyaltyRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   2775
         Width           =   1395
      End
      Begin VB.Frame Frame3 
         Caption         =   "BIC"
         ForeColor       =   &H8000000D&
         Height          =   1695
         Left            =   -74640
         TabIndex        =   92
         Top             =   1650
         Width           =   3720
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
            TabIndex        =   93
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
            TabIndex        =   91
            Top             =   720
            Width           =   3360
         End
      End
      Begin VB.TextBox txtCategoryHeading 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   330
         Left            =   -74700
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   3480
         Width           =   10770
      End
      Begin VB.TextBox txtFlagText 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -74730
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   2505
         Width           =   7215
      End
      Begin VB.TextBox txtReturnable 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   3060
         Width           =   450
      End
      Begin VB.TextBox txtProductType 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73635
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   450
         Width           =   2745
      End
      Begin MSComCtl2.DTPicker DTMMSince 
         Height          =   390
         Left            =   9825
         TabIndex        =   78
         Top             =   4020
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
         Format          =   49872897
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
         TabIndex        =   77
         Top             =   4020
         Width           =   1725
      End
      Begin VB.TextBox txtDefaultDeliveryDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -65490
         Locked          =   -1  'True
         TabIndex        =   75
         Top             =   3540
         Width           =   1395
      End
      Begin VB.TextBox txtSS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -73140
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   3525
         Width           =   1395
      End
      Begin VB.TextBox txtLastCountedPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -70185
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2805
         Width           =   960
      End
      Begin VB.TextBox txtLastCountedQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -70785
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   2805
         Width           =   555
      End
      Begin VB.TextBox txtLastCounted 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -72690
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2805
         Width           =   1860
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   2040
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   4050
         Width           =   3705
      End
      Begin VB.TextBox txtPublisher 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1210
         Width           =   4980
      End
      Begin VB.TextBox txtEdition 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1600
         Width           =   4980
      End
      Begin VB.TextBox txtPubDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   820
         Width           =   4980
      End
      Begin VB.TextBox txtPubPlace 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   285
         Left            =   -69090
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   430
         Width           =   4980
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
         Left            =   3255
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2040
         Width           =   3705
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   3120
         Width           =   2550
      End
      Begin VB.Frame Frame2 
         Caption         =   "Aged stock"
         ForeColor       =   &H8000000D&
         Height          =   795
         Left            =   -74355
         TabIndex        =   36
         Top             =   480
         Width           =   5220
         Begin VB.TextBox txt18Plus 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt18 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   3270
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   2415
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txt6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1545
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   330
            Width           =   825
         End
         Begin VB.TextBox txtAgedDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   195
            Locked          =   -1  'True
            TabIndex        =   37
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
         TabIndex        =   33
         Top             =   1590
         Width           =   7230
      End
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         Height          =   510
         Left            =   -74715
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   690
         Width           =   7230
      End
      Begin VB.TextBox txtUSPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   3405
         Width           =   855
      End
      Begin VB.TextBox txtUKPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3075
         Width           =   855
      End
      Begin VB.TextBox txtTotalSold 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   750
      End
      Begin VB.TextBox txtCost 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1920
         Width           =   1170
      End
      Begin VB.TextBox txtRRP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   1170
      End
      Begin VB.TextBox txtReserved 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H000040C0&
         Height          =   285
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   600
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
         TabIndex        =   15
         Top             =   600
         Width           =   750
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         ForeColor       =   &H8000000D&
         Height          =   285
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -65490
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3150
         Width           =   1395
      End
      Begin TrueOleDBGrid60.TDBGrid POGrid 
         Height          =   1590
         Left            =   3090
         OleObjectBlob   =   "frmProductPrev2.frx":6CE3
         TabIndex        =   51
         Top             =   420
         Width           =   4020
      End
      Begin TrueOleDBGrid60.TDBGrid COGrid 
         Height          =   1335
         Left            =   3090
         OleObjectBlob   =   "frmProductPrev2.frx":B636
         TabIndex        =   62
         Top             =   2655
         Width           =   4005
      End
      Begin TrueOleDBGrid60.TDBGrid APPGRID 
         Height          =   1575
         Left            =   7395
         OleObjectBlob   =   "frmProductPrev2.frx":FAD1
         TabIndex        =   64
         Top             =   435
         Width           =   3720
      End
      Begin TrueOleDBGrid60.TDBGrid MMGRID 
         Height          =   1335
         Left            =   7395
         OleObjectBlob   =   "frmProductPrev2.frx":13F6D
         TabIndex        =   69
         Top             =   2640
         Width           =   3735
      End
      Begin TrueOleDBGrid60.TDBGrid GSB 
         Height          =   1320
         Left            =   -73410
         OleObjectBlob   =   "frmProductPrev2.frx":18070
         TabIndex        =   105
         Top             =   2550
         Width           =   9495
      End
      Begin TrueOleDBGrid60.TDBGrid GSF 
         Height          =   1500
         Left            =   -73410
         OleObjectBlob   =   "frmProductPrev2.frx":1C06C
         TabIndex        =   106
         Top             =   960
         Width           =   9495
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   2790
         Left            =   -74790
         OleObjectBlob   =   "frmProductPrev2.frx":20198
         TabIndex        =   109
         Top             =   660
         Width           =   10965
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "System notes"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -67185
         TabIndex        =   142
         Top             =   2265
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EUR"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   105
         TabIndex        =   116
         Top             =   3765
         Width           =   465
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "blue indicates data updated in dayend only"
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
         Height          =   285
         Left            =   -68400
         TabIndex        =   112
         Top             =   870
         Width           =   3855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFC0&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   -68880
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label43 
         BackStyle       =   0  'Transparent
         Caption         =   "Possible substitutions for this item"
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   -74760
         TabIndex        =   108
         Top             =   2970
         Width           =   1185
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "This item can substitute for these . . ."
         ForeColor       =   &H8000000D&
         Height          =   555
         Left            =   -74760
         TabIndex        =   107
         Top             =   1500
         Width           =   1185
      End
      Begin VB.Label lblNDA 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No discount allowed"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   180
         TabIndex        =   102
         Top             =   2790
         Visible         =   0   'False
         Width           =   2790
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Special"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1815
         TabIndex        =   101
         Top             =   1620
         Width           =   930
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Catalogues"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -67185
         TabIndex        =   99
         Top             =   450
         Width           =   1035
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Loyalty Rate"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -70905
         TabIndex        =   95
         Top             =   2805
         Width           =   1770
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Catalogue heading"
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -74790
         TabIndex        =   90
         Top             =   3165
         Width           =   1680
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Flag text"
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   -74670
         TabIndex        =   88
         Top             =   2205
         Width           =   705
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Returnable"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1395
         TabIndex        =   86
         Top             =   3090
         Width           =   975
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   -70080
         TabIndex        =   84
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label37 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   -70935
         TabIndex        =   83
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label36 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H8000000D&
         Height          =   210
         Left            =   -72195
         TabIndex        =   82
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Product type"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74865
         TabIndex        =   81
         Top             =   465
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000C0&
         Height          =   420
         Left            =   135
         TabIndex        =   79
         Top             =   4065
         Width           =   1905
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lead time"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -67110
         TabIndex        =   76
         Top             =   3570
         Width           =   1590
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Distributor"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   -70905
         TabIndex        =   74
         Top             =   2055
         Width           =   1770
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last deal"
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   -70905
         TabIndex        =   73
         Top             =   2445
         Width           =   1770
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Order by seesafe"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74895
         TabIndex        =   71
         Top             =   3570
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last counted"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74220
         TabIndex        =   68
         Top             =   2835
         Width           =   1395
      End
      Begin VB.Label lblPublisher 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -70905
         TabIndex        =   60
         Top             =   1260
         Width           =   1770
      End
      Begin VB.Label lblEdition 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Edition"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -70905
         TabIndex        =   104
         Top             =   1665
         Width           =   1770
      End
      Begin VB.Label lblPublicationDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Publication date"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -70905
         TabIndex        =   58
         Top             =   855
         Width           =   1770
      End
      Begin VB.Label lblPublicationPlace 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Publication place"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -70905
         TabIndex        =   57
         Top             =   450
         Width           =   1770
      End
      Begin VB.Label lblSupplier 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   -69090
         TabIndex        =   47
         Top             =   1995
         Width           =   3135
      End
      Begin VB.Label lblDeal 
         Appearance      =   0  'Flat
         BackColor       =   &H00DBFAFB&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -69090
         TabIndex        =   46
         Top             =   2385
         Width           =   2295
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Last modified"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -67905
         TabIndex        =   45
         Top             =   3585
         Width           =   1395
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Added"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -67815
         TabIndex        =   44
         Top             =   3195
         Width           =   1290
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comment"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -74685
         TabIndex        =   35
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -74715
         TabIndex        =   34
         Top             =   390
         Width           =   1035
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "USD"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   105
         TabIndex        =   31
         Top             =   3420
         Width           =   465
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GBP"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   3075
         Width           =   450
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total sold"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2040
         TabIndex        =   29
         Top             =   390
         Width           =   795
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "weighted avg. cost (Ex VAT)"
         ForeColor       =   &H8000000D&
         Height          =   405
         Left            =   30
         TabIndex        =   28
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R.R.P."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   465
         TabIndex        =   26
         Top             =   975
         Width           =   540
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reserved"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1050
         TabIndex        =   25
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "On hand"
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   180
         TabIndex        =   24
         Top             =   390
         Width           =   705
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Categories"
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
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   -66600
         TabIndex        =   9
         Top             =   3180
         Width           =   1080
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Sell.P."
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1905
         TabIndex        =   27
         Top             =   975
         Width           =   540
      End
   End
   Begin VB.TextBox txtSubtitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   4560
   End
   Begin VB.TextBox txtAuthor 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   105
      Width           =   4155
   End
   Begin VB.TextBox txtTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00DBFAFB&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   465
      Width           =   4560
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Ex VAT"
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   8955
      TabIndex        =   139
      Top             =   555
      Width           =   1050
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "RRP"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8955
      TabIndex        =   138
      Top             =   255
      Width           =   765
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last sold"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5595
      TabIndex        =   137
      Top             =   855
      Width           =   870
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   8085
      TabIndex        =   136
      Top             =   15
      Width           =   840
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Qty"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   7350
      TabIndex        =   135
      Top             =   15
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   6555
      TabIndex        =   134
      Top             =   15
      Width           =   840
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last rec'd"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5400
      TabIndex        =   133
      Top             =   555
      Width           =   1080
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Last ord'd"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   5505
      TabIndex        =   132
      Top             =   240
      Width           =   975
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
      Left            =   4710
      TabIndex        =   97
      Top             =   6075
      Width           =   2940
   End
   Begin VB.Label lblServiceItem 
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
      Left            =   3615
      TabIndex        =   72
      Top             =   6615
      Width           =   1785
   End
   Begin VB.Label lblAuthor 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   -15
      TabIndex        =   5
      Top             =   150
      Width           =   1095
   End
   Begin VB.Label lblSubTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtitle"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   870
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   -45
      TabIndex        =   3
      Top             =   510
      Width           =   1125
   End
End
Attribute VB_Name = "frmProductPrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngID As Long
Dim strPID As String
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
Dim lngSTGridHeight As Long
Dim lngAPPGRIDHeight As Long
Dim lngMMGRIDHeight As Long
Dim lngPOGridHeight As Long
Dim lngCOGridHeight As Long


Public Property Get Timing() As String
    On Error GoTo errHandler
    Timing = strTime
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Timing"
End Property
Sub component(pProduct As a_Product, Optional pstrTime As String)
    On Error GoTo errHandler
    strTime = pstrTime
    strTime = strTime & "Start frmProductPrev:component:" & Now() & vbCrLf
    Set oProd = Nothing
    Set oProd = pProduct

    Set tlCatHead = Nothing
    Set tlCatHead = New z_TextList
    tlCatHead.Load ltCatalogueHeadings



strTime = strTime & "End frmProductPrev:component:" & Now() & vbCrLf
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProductPrev.Component(pProduct,pstrTime)", Array(pProduct, pstrTime)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.component(pProduct,pstrTime)", Array(pProduct, pstrTime)
End Sub
Public Property Get CategoryName() As String
    On Error GoTo errHandler
    CategoryName = tlSections.Item(oProd.CategoryID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.CategoryName"
End Property
Public Property Get ProductTypeName() As String
    On Error GoTo errHandler
    ProductTypeName = tlProductTypes.Item(oProd.ProductTypeID)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.ProductTypeName"
End Property



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
    ErrorIn "frmProductPrev.cmdBack_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDefaultColWidths_Click()
    On Error GoTo errHandler
    Me.MMGRID.Columns(0).Width = 900
    Me.MMGRID.Columns(1).Width = 900
    Me.MMGRID.Columns(2).Width = 500
    Me.MMGRID.Columns(3).Width = 500
    
    Me.APPGRID.Columns(0).Width = 900
    Me.APPGRID.Columns(1).Width = 900
    Me.APPGRID.Columns(2).Width = 500
    Me.APPGRID.Columns(3).Width = 500
    Me.APPGRID.Columns(4).Width = 500
    
    Me.COGrid.Columns(0).Width = 900
    Me.COGrid.Columns(1).Width = 900
    Me.COGrid.Columns(2).Width = 500
    Me.COGrid.Columns(3).Width = 500
    Me.COGrid.Columns(4).Width = 500

    Me.POGrid.Columns(0).Width = 900
    Me.POGrid.Columns(1).Width = 900
    Me.POGrid.Columns(2).Width = 500
    Me.POGrid.Columns(3).Width = 500
    Me.POGrid.Columns(4).Width = 500

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdDefaultColWidths_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdExplainCost_Click()
    On Error GoTo errHandler
Dim f As frmExplainCost
Dim x As Long
Dim Y As Long
Dim lRS As ADODB.Recordset
Dim OpenResult As Integer
 
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    Set f = New frmExplainCost

    x = MouseX(Me.hWnd)
    Y = MouseY(Me.hWnd)
    If Me.Width / Screen.TwipsPerPixelX / 2 > x Then
        x = x + 11
    Else
        x = x - (f.Width / Screen.TwipsPerPixelX) - 30 '- 200
    End If
    If Me.Height / Screen.TwipsPerPixelY / 2 > Y Then
        Y = Y + 11
    Else
        Y = Y - (f.Height / Screen.TwipsPerPixelY) + 100
    End If
    Set lRS = New ADODB.Recordset
    lRS.CursorLocation = adUseClient
    lRS.Open "SELECT * FROM vExplainAvgCost WHERE PID = '" & oProd.PID & "' ORDER BY dte", oPC.COShort
    f.component lRS
    f.TOP = Y * Screen.TwipsPerPixelY
    f.Left = x * Screen.TwipsPerPixelX
    f.component lRS
    f.Show vbModal
   
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdExplainCost_Click", , EA_NORERAISE
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
    ErrorIn "frmProductPrev.cmdForward_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRecon_Click()
    On Error GoTo errHandler
Dim frmMM As frmMovements
    oProd.ReloadRecentMovements
    Set frmMM = New frmMovements
    frmMM.component oProd, Me.TOP + 200, Me.Left + 1000
    frmMM.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdRecon_Click", , EA_NORERAISE
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
    ErrorIn "frmProductPrev.cmdSales_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdStatusChange_Click()
    On Error GoTo errHandler
Dim frm As New frmPreDeliveryAdvice
    
    frm.component "", oProd.PID, "R", oProd.StatusF
    frm.Show vbModal
    If frm.GetNewStatus > "" Then
        Me.lblStatus = frm.GetNewStatus
    End If
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdStatusChange_Click", , EA_NORERAISE
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
    ErrorIn "frmProductPrev.cmdSub_Click", , EA_NORERAISE
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
    ErrorIn "frmProductPrev.cmdSubs_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdWsSales_Click()
    On Error GoTo errHandler
Dim frm As New frmSalesWS
    frm.component oProd
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdWsSales_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub COGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If InStr(1, XD(Bookmark, 8), "*") > 0 Then
        RowStyle.Font.Bold = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.COGrid_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub Command1_Click()
    On Error GoTo errHandler
  '  ErrorIn "TEST ERROR CRASH for CLEANUP"
'  oProd.Crash
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Command2_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Command2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub SetMenu()
    On Error GoTo errHandler
    Forms(0).mnuAdjust.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.SetMenu"
End Sub
Private Sub COGrid_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    On Error GoTo errHandler
    mnuSaveLayout
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.COGrid_ColResize(ColIndex,Cancel)", Array(ColIndex, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub APPGRID_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    On Error GoTo errHandler
    mnuSaveLayout
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.APPGRID_ColResize(ColIndex,Cancel)", Array(ColIndex, Cancel), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub MMGRID_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    On Error GoTo errHandler
    mnuSaveLayout
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.MMGRID_ColResize(ColIndex,Cancel)", Array(ColIndex, Cancel), EA_NORERAISE
    HandleError
End Sub


Private Sub POGrid_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    On Error GoTo errHandler
    mnuSaveLayout
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.POGrid_ColResize(ColIndex,Cancel)", Array(ColIndex, Cancel), EA_NORERAISE
    HandleError
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout POGrid, "POGrid"
    SaveLayout COGrid, "COGrid"
    SaveLayout APPGRID, "APPGrid"
    SaveLayout MMGRID, "MMGrid"
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn Me.Name & ":mnuSaveLayout", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.mnuSaveLayout"
End Sub

Private Sub cboCatHead_Click()
    On Error GoTo errHandler
'    oProd.setCatalogueheadingID tlCatHead.Key(cboCatHead)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cboCatHead_Click", , EA_NORERAISE
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
    ErrorIn "frmProductPrev.cmdDelete_Click", , EA_NORERAISE
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
    ErrorIn "frmProductPrev.APPGRID_DblClick", , EA_NORERAISE
    HandleError
End Sub
'
'Private Sub chart1_AxisLabelSelected(axisID As Integer, AxisIndex As Integer, labelSetIndex As Integer, LabelIndex As Integer, MouseFlags As Integer, Cancel As Integer)
'Cancel = True
'End Sub
'
'Private Sub chart1_AxisTitleSelected(axisID As Integer, AxisIndex As Integer, MouseFlags As Integer, Cancel As Integer)
'Cancel = True
'End Sub
'
'Private Sub chart1_ChartSelected(MouseFlags As Integer, Cancel As Integer)
'Cancel = True
'End Sub
'
'Private Sub chart1_FootnoteSelected(MouseFlags As Integer, Cancel As Integer)
'Cancel = True
'End Sub
'
'Private Sub chart1_LegendSelected(MouseFlags As Integer, Cancel As Integer)
'Cancel = True
'End Sub
'
'Private Sub chart1_PlotSelected(MouseFlags As Integer, Cancel As Integer)
'Cancel = True
'End Sub

'Private Sub chart1_PointSelected(Series As Integer, DataPoint As Integer, MouseFlags As Integer, Cancel As Integer)
''MsgBox "Series " & Series
''MsgBox "Datapoint " & DataPoint
'    Select Case Series
'    Case 1
'        MsgBox "Sales qty = " & oProd.CurrentSales.FindByWeek(DataPoint).qty & vbCrLf & "Sales value = " & oProd.CurrentSales.FindByWeek(DataPoint).ValuF
'    Case 2
'   '     MsgBox "Sales qty = " & oProd.PreviousSales.FindByWeek(DataPoint).Qty & vbCrLf & "Sales value = " & oProd.PreviousSales.FindByWeek(DataPoint).ValuF
'    End Select
'End Sub
'
'Private Sub chart1_SeriesSelected(Series As Integer, MouseFlags As Integer, Cancel As Integer)
''Cancel = True
'End Sub

Private Sub cmdALLMM_Click()
    On Error GoTo errHandler
Dim frmMM As frmMovements_All
    oProd.ReloadMovements Me.DTMMSince
    Set frmMM = New frmMovements_All
    frmMM.component oProd
    frmMM.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdALLMM_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmProduct
    Set frm = New frmProduct
    frm.component oProd, Me
    frm.Show
    Exit Sub
Errh:
    MsgBox Error

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdEdit_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub formatgrid()
    On Error GoTo errHandler
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.formatgrid"
End Sub
Private Sub cmdsearchisbn_Click()
    On Error GoTo errHandler
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
        Me.Caption = "Product master preview -  Stock code: " & .code
        txtSubtitle = .SubTitle
        txtTitle = .Title
        txtPublisher = .Publisher
        
    End With
    LoadControls
    LoadCopies
    LoadStock
    LoadMovements
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdsearchisbn_Click", , EA_NORERAISE
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
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.COGrid_DblClick", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Set oProd = Nothing
    Set tlCatHead = Nothing
    Set tlSections = Nothing
    Set tlProductTypes = Nothing
    Set XA = Nothing
    Set XB = Nothing
    Set XC = Nothing
    Set XD = Nothing
    Set XE = Nothing
    Set XF = Nothing
    Set XSF = Nothing
    Set XSB = Nothing
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo errHandler
'Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Grid1_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, KeyAscii, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
            Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub


Private Sub GSB_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(GSB.Bookmark) Then Exit Sub
    If XSB.UpperBound(1) = 0 Then Exit Sub
    str = FNS(XSB.Value(GSB.Bookmark, 1))
    If str = "" Then Exit Sub
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.GSB_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub GSF_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(GSF.Bookmark) Then Exit Sub
    If XSF.UpperBound(1) = 0 Then Exit Sub
    str = FNS(XSF.Value(GSF.Bookmark, 1))
    If str = "" Then Exit Sub
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.GSF_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub GSF_DblClick()
    On Error GoTo errHandler
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
        frmA.component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.GSF_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub GSB_DblClick()
    On Error GoTo errHandler
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
        frmA.component oP
        frmA.Show
    Else
        Set frm = New frmProductPrev
        frm.component oP
        frm.Show
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.GSB_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub MMGRID_DblClick()
    On Error GoTo errHandler
Dim strType As String
Dim frm As Form
Dim i As Integer
Dim dteDocDate, dteLimitToView As Date
Dim tmpType As String

    If IsNull(MMGRID.Bookmark) Then
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass

    strType = Trim(XF(MMGRID.Bookmark, 4))
    If (InStr(1, strType, "(") > 0) Then
        strType = Left(strType, InStr(1, strType, "(") - 1)
    End If
    Select Case strType  '' IIf(InStr(1, strType, "(") > 0, Left(strType, InStr(1, strType, "(") - 1), strType)
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
    Case "CS", "POS"
        If oPC.BlindCashup = True Then
            Dim oSQL As New z_SQL
            dteDocDate = XF.Value(MMGRID.Bookmark, 8)
            dteLimitToView = oSQL.GetDateOfEarliestUnSignedSession
            If dteDocDate >= StartOfDay(dteLimitToView) Then
                MsgBox "There are unsigned cash ups starting prior to your selected end date (" & Format(dteLimitToView, "dd/mm/yyyy") & "). You cannot include thse in the report. Select an earlier end date.", vbInformation, "Can't do this"
                Exit Sub
            End If
        End If
    
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
    Case "ADJ"
        MsgBox XF(MMGRID.Bookmark, 1), vbInformation, "ad-hoc adjustment"
    End Select
    Screen.MousePointer = vbDefault

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.MMGRID_DblClick", , EA_NORERAISE
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
    ErrorIn "frmProductPrev.mnuAdjust"
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
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.POGrid_DblClick", , EA_NORERAISE
    HandleError
End Sub



Private Sub StGrid_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo errHandler
'Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.StGrid_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, KeyAscii, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub StGrid_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
    Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.StGrid_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
'MsgBox "Selected row is : " & Grid1.Row + 1
Dim frm As frmCopyPreview
Dim oCopy As a_Copy
    If IsNull(Grid1.Bookmark) Then Exit Sub
    Set oCopy = oProd.Copies(XA(Grid1.Bookmark, 7))
    Set frm = New frmCopyPreview

    frm.component oCopy, oProd
    frm.Show ' vbModal

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Grid1_DblClick", , EA_NORERAISE
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11950
        Height = 6800
    End If
    Set tlSections = oPC.Configuration.Sections  'New z_TextList
    SSTab1.TabVisible(5) = oPC.Configuration.AllowCopyInfo
    Set tlProductTypes = oPC.Configuration.ProductTypes   'New z_TextList
    LoadControls
    SSTab1.TabCaption(5) = "&6.Copies (" & oProd.Copies.CountForSale & ")"
    SSTab1.Tab = 0
    Me.cmdWsSales.Visible = oPC.ShowWordstockSales
    SetGridLayout POGrid, "POGrid"
    SetGridLayout COGrid, "COGrid"
    SetGridLayout APPGRID, "APPGRID"
    SetGridLayout MMGRID, "MMGRID"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Form_Load", , EA_NORERAISE
    HandleError
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
    ErrorIn "frmProductPrev.LoadProductSections"
End Sub

Public Sub RefreshForm()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.RefreshForm"
End Sub
Private Sub LoadControls()
10        On Error GoTo errHandler
      Dim errRepeat As Integer

20        errSysHandlerSet
30        flgLoading = True
40      On Error Resume Next
50        cmdWash.Visible = oPC.BFLoaded
60      On Error GoTo errHandler
70        Me.Caption = "Product master preview -  Stock code: " & oProd.code & "      EAN: " & oProd.EAN
80        Me.txtProductType = tlProductTypes.Item(oProd.ProductTypeID)
90        txtTitle = oProd.Title
100       txtSubtitle = oProd.SubTitle
110       txtAuthor = oProd.Author
120       txtEdition = oProd.Edition
130       txtPublisher = oProd.Publisher
140       txtPubPlace = oProd.PublicationPlace
150       Me.txtPubDate = oProd.PublicationDate
160       Me.txtPubDate = oProd.PublicationDate
170       Me.txtBinding = oProd.BindingCode
        '  Me.txtSection = tlProductTypes.Item(oProd.CategoryID)
180       txtDefaultDeliveryDays = oProd.DefaultDeliveryDays
190       Me.txtCategoryHeading = tlCatHead.Item(oProd.CatalogueheadingID)
200       Me.txtRRP = oProd.RRPF
210       If oProd.SpecialPrice > 0 Then
220     txtSP.BackColor = &HDBFAFB
230     txtSP.FontSize = 9
240     txtSP.FontBold = False
250     txtSP.ForeColor = vbGrayText
260     txtSSP.BackColor = vbYellow
270     txtSSP.FontSize = 10
280     txtSSP.FontBold = True
290     txtSSP.ForeColor = &H8000000D
300     txtSP = "(" & oProd.SPF & ")"
310     txtSSP = oProd.SpecialPriceF
320       Else
330     txtSP.BackColor = vbYellow
340     txtSP.FontSize = 10
350     txtSP.FontBold = True
360     txtSP.ForeColor = &H8000000D
370     txtSSP.BackColor = &HDBFAFB
380     txtSSP.FontSize = 9
390     txtSSP.FontBold = False
400     txtSSP.ForeColor = vbGrayText
410     txtSP = oProd.SPF
420     txtSSP = oProd.SpecialPriceF
430       End If
440       Me.txtCost = oProd.CostF
450       Me.txtTotalSold = oProd.QtyTotalSold
460       Me.txtUKPrice = oProd.UKPriceF
470       Me.txtUSPrice = oProd.USPriceF
480       Me.txtEUPrice = oProd.EUPriceF
490       Me.txtWeight = oProd.WeightF
500       Me.txtAgedDate = ""
510       Me.txtComment = oProd.Comment
520       Me.txtNotes = oProd.Note
530       Me.txtCost = oProd.CostF
540       Me.txtDateAdded = oProd.DateRecordAddedF
550       Me.txtDateLastModified = oProd.DateLastModifiedF
560       Me.txtDescription = oProd.Description
570       Me.txtLastCounted = oProd.dateLastCountedF
580       Me.txtLastCountedPrice = oProd.PriceLastCountedF
590       Me.txtLastCountedQty = oProd.QtyLastCountedF
600       Me.txtLastReceived = oProd.DateLastDeliveredF
610       Me.txtLastReceivedPrice = oProd.PriceLastDeliveredF
620       Me.txtLastReceivedQty = oProd.QtyLastDeliveredF
630       Me.txtLastOrdered = oProd.DateLastOrderedF
640       Me.txtLastOrderedPrice = oProd.PriceLastOrderedF
650       Me.txtLastOrderedQty = oProd.QtyLastOrderedF
660       Me.txtLastSoldDate = oProd.DateLastSoldF
670       Me.txtLastSoldPrice = oProd.PriceLastSoldF
680       Me.txtLastSoldQty = oProd.QtylastSold
690       Me.txtOnHand = oProd.QtyOnHandF
700       Me.txtCatalogues = oProd.CatalogueEntries_Concat
          'If multibuy there is automatically not discount allowed
710       If oProd.MultibuyCode > "" Then
720     lblNDA.Visible = True
730     lblNDA.Caption = oPC.Configuration.Multibuys.ItemByF4(oProd.MultibuyCode)
740       Else
750     lblNDA.Visible = oProd.IsNDA
760       End If
770       Me.chkCore = IIf(oProd.IsCore, 1, 0)
      '    Me.txtSummary = oProd.Summary
780       Me.txtVAT = oProd.VATRateToUseF
790       Me.txtReserved = oProd.QtyReservedF
800       Me.txtReturnable = oProd.ReturnAvailability
810       txtFlagText = oProd.FlagText
820       txtBIC = oProd.BIC
830       txtLoyaltyRate = oProd.LoyaltyRateF
840       txtBICDescription = oPC.Configuration.BICs.FetchBICDescriptionsFromCodeSet(txtBIC)
      '    txtBICDescription = oProd.BICDescription
850       Me.lblServiceITem.Visible = oProd.IsServiceItem
860       Me.lblObsolete = IIf(oProd.Obsolete, "Obsolete", "")
         ' MsgBox oProd.Status
870       txtSS = IIf(oProd.Seesafe = True, "Yes", "No")
880       Me.lblSupplier.Caption = oProd.LastSupplierName
890       Me.lblDeal.Caption = oProd.LastDealDescription
900       lblStatus = oProd.StatusF
910       DTMMSince.Value = IIf(oProd.DateLastCounted < DTMMSince.MinDate, DTMMSince.MinDate, DateAdd("yyyy", -1, oProd.DateLastCounted))
920       LoadMovements
930       LoadCopies
940       LoadStock
950       LoadProductSections
960       flgLoading = False
970       Exit Sub
errHandler:
980       ErrPreserve
990       If Err.Number = -2147217407 Then   'Access violation
1000          errRepeat = errRepeat + 1
1010          LogSaveToFile "Access violation in frmProductPrev: LoadControls, err repeat = " & CStr(errRepeat) & ", line:" & CStr(Erl())
1020          If errRepeat < 5 Then
1030              Resume Next
1040          Else
1050              LogSaveToFile "Access violation in frmProductPrev: LoadControls after 5 re-attempts"
1060              MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't run search."
1070              Err.Clear
1080              Exit Sub
1090          End If
1100      End If

1110      If ErrMustStop Then Debug.Assert False: Resume
1120      ErrorIn "frmProductPrev.LoadControls"
End Sub
Private Sub LoadCopies()
    On Error GoTo errHandler
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.LoadCopies"
End Sub
Private Sub LoadStock()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

'    Set XB = New XArrayDB
'    XB.Clear
'    XB.ReDim 1, oProd.Stores.Count, 1, 6
'    lngIndex = 0
'    Do While lngIndex < oProd.Stores.Count
'        lngIndex = lngIndex + 1
'        XB.Value(lngIndex, 1) = oProd.Stores(lngIndex).Storename & oProd.Stores(lngIndex).LastSharedDateFShortwithParentheses
'        XB.Value(lngIndex, 2) = oProd.Stores(lngIndex).QtyOnHand
'        XB.Value(lngIndex, 3) = oProd.Stores(lngIndex).QtyReserved
'        XB.Value(lngIndex, 4) = oProd.Stores(lngIndex).QtyOnBackorder
'        XB.Value(lngIndex, 5) = oProd.Stores(lngIndex).QtyOnOrder
'        XB.Value(lngIndex, 6) = oProd.Stores(lngIndex).QtyCopiesOnHand
'    Loop
'    XB.QuickSort 1, oProd.Stores.Count, 1, XORDER_ASCEND, XTYPE_STRING
'    StGrid.Array = XB
'    StGrid.ReBind
    If oPC.IsMultiStore Then
        StGrid.Visible = True
            
        Set XB = New XArrayDB
        XB.Clear
        XB.ReDim 1, oProd.Stores.Count, 1, 6
        lngIndex = 0
        Do While lngIndex < oProd.Stores.Count
            lngIndex = lngIndex + 1
            XB.Value(lngIndex, 1) = oProd.Stores(lngIndex).StoreCode '& oProd.Stores(lngIndex).LastSharedDateFShortwithParentheses
            XB.Value(lngIndex, 2) = oProd.Stores(lngIndex).QtyOnHand
            XB.Value(lngIndex, 3) = oProd.Stores(lngIndex).QtyReserved
            XB.Value(lngIndex, 4) = oProd.Stores(lngIndex).QtyOnBackorder
            XB.Value(lngIndex, 5) = oProd.Stores(lngIndex).QtyonOrder
   '         XB.Value(lngIndex, 6) = oProd.Stores(lngIndex).QtyCopiesOnHand
        Loop
        XB.QuickSort 1, oProd.Stores.Count, 1, XORDER_ASCEND, XTYPE_STRING
        StGrid.Array = XB
        StGrid.ReBind
    Else
        StGrid.Visible = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.LoadStock"
End Sub
Private Sub lvwCopies_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.lvwCopies_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
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
        MMGRID.Splits(0).Caption = oProd.dateLastCountedF & ": " & " (" & oProd.dateLastCountedF & ":" & oProd.QtyLastCountedF & ":" & oProd.CostAtLastStockTakeF & ")"
    End If
    
    txtTotalOSPO = "We are awaiting " & lngQtyOutstanding & " copies."
    txtTotalOSCO = "Customers are awaiting " & lngQtySpecial & " copies."
    txtTotalOSAP = "Expecting return of " & lngQtyApps & " copies."
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.LoadMovements"
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
    ErrorIn "frmProductPrev.LoadPOs"
End Sub
Private Sub LoadCOs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XD = New XArrayDB
    XD.Clear
    XD.ReDim 1, oProd.OSCOs.Count, 1, 8
    For lngIndex = 1 To oProd.OSCOs.Count
        XD.Value(lngIndex, 1) = oProd.OSCOs(lngIndex).DOCCode
        XD.Value(lngIndex, 2) = oProd.OSCOs(lngIndex).DocDateF
        XD.Value(lngIndex, 3) = oProd.OSCOs(lngIndex).TPNAME
        XD.Value(lngIndex, 4) = oProd.OSCOs(lngIndex).COLQty
        XD.Value(lngIndex, 5) = oProd.OSCOs(lngIndex).COLCollected
        XD.Value(lngIndex, 6) = oProd.OSCOs(lngIndex).TRID
        XD.Value(lngIndex, 7) = oProd.OSCOs(lngIndex).DateForSort
        XD.Value(lngIndex, 8) = oProd.OSCOs(lngIndex).Ref
    Next
    XD.QuickSort 1, oProd.OSCOs.Count, 7, XORDER_DESCEND, XTYPE_STRING
    COGrid.Array = XD
    COGrid.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.LoadCOs"
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
        XE.Value(lngIndex, 7) = oProd.OSAPs(lngIndex).DOCDate
    Next
    XE.QuickSort 1, oProd.OSAPs.Count, 7, XORDER_DESCEND, XTYPE_DATE
    APPGRID.Array = XE
    APPGRID.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.LoadAPs"
End Sub
Private Sub LoadMMs()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String

    Set XF = New XArrayDB
    XF.Clear
    XF.ReDim 1, oProd.MMs.Count, 1, 8
    For lngIndex = 1 To oProd.MMs.Count
        XF.Value(lngIndex, 1) = oProd.MMs(lngIndex).DOCCode
        XF.Value(lngIndex, 2) = oProd.MMs(lngIndex).DocDateF
        XF.Value(lngIndex, 3) = oProd.MMs(lngIndex).Qty
    '    XF.Value(lngIndex, 4) = oProd.MMs(lngIndex).typ
        XF.Value(lngIndex, 4) = oProd.MMs(lngIndex).typ & IIf(oProd.MMs(lngIndex).typ = "POS", "(" & oProd.MMs(lngIndex).Station & ")", "")
        XF.Value(lngIndex, 5) = oProd.MMs(lngIndex).PID
        XF.Value(lngIndex, 6) = oProd.MMs(lngIndex).TRID
        XF.Value(lngIndex, 7) = oProd.MMs(lngIndex).Seq
        XF.Value(lngIndex, 8) = oProd.MMs(lngIndex).DOCDate
    Next
    XF.QuickSort 1, oProd.MMs.Count, 7, XORDER_DESCEND, XTYPE_INTEGER
    MMGRID.Array = XF
    MMGRID.ReBind

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.LoadMMs"
End Sub

Private Sub Form_DblClick()
    On Error GoTo errHandler
    If Not IsNull(oProd) Then
    
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText oProd.ProductDetails
    End If
    TouchRecord
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.Form_DblClick", , EA_NORERAISE
    HandleError
End Sub

Public Sub ExportInCatalogueFormat()
    On Error GoTo errHandler
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
'Dim objXSL As New MSXML2.DOMDocument60
'Dim opXMLDOC As New MSXML2.DOMDocument60
'Dim objXMLDOC  As New MSXML2.DOMDocument60
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

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.ExportInCatalogueFormat"
End Sub
Public Sub mnuTouchRecord()
    On Error GoTo errHandler
    TouchRecord
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.mnuTouchRecord"
End Sub
Public Sub TouchRecord()
    On Error GoTo errHandler
Dim oSQL As New z_SQL

    oSQL.RunSQL "INSERT INTO tPRODUPDATES(PRU_LOG_TYPE,PRU_P_ID,PRU_Code,PRU_EAN," _
            & "PRU_Publisher,PRU_SeriesTitle,PRU_MainAuthor,PRU_Title,PRU_SP,PRU_VATRATE,PRU_LoyaltyRATE," _
            & "PRU_PTID,PRU_SECID) " _
            & "SELECT 'NEW',P_ID,P_CODE," & "P_EAN,P_PUBLISHER,P_SERIESTITLE,P_MAINAUTHOR," _
            & "P_TITLE,P_SP,dbo.VATRATETOUSE(P_SpecialVat,P_VatRate),P_LoyaltyRATE, P_ProductType_ID, vSectionMaster.PSEC_SEC_ID " _
            & " FROM tPRODUCT LEFT JOIN vSectionMaster ON P_ID = vSectionMaster.PSEC_P_ID " _
            & " WHERE P_ID = '" & oProd.PID & "'"

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.TouchRecord"
End Sub
Private Sub cmdWash_Click()
Dim Res As Boolean

    On Error GoTo errHandler
    If MsgBox("This operation will overwrite title, author, publisher etc. with the data on Nielsen. Confirm", vbExclamation + vbOKCancel, "Warniong") = vbCancel Then
        Exit Sub
    End If
    strPID = oProd.PID
    UpdateFromBookfind True, True, True, False, True, True, True, True, True, True, True, True, True, False, 0, oProd.EAN
    Set oProd = Nothing
    Set oProd = New a_Product
    Res = oProd.Load(strPID, 0)
    If Res = 0 Then   'item successfully retrieved
        LoadControls
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdWash_Click", , EA_NORERAISE
    HandleError
End Sub

'=============This code is a copy of code in z_Batch
Function UpdateFromBookfind(pAuthor As Boolean, _
pTitle As Boolean, _
pSubtitle As Boolean, _
pAvailability As Boolean, _
pBindingcode As Boolean, _
pEdition As Boolean, _
pSUpplierCode As Boolean, _
pPublisherName As Boolean, _
pSeriesTitle As Boolean, _
pPublicationDate As Boolean, _
pUKPrice As Boolean, _
pRRP As Boolean, _
pBIC As Boolean, _
pBookStatus As Boolean, _
Optional pSTAFFID As Long, _
Optional pEAN As String)
    On Error GoTo errHandler
Dim xml As String
Dim mudtProps As ProductProps
Dim strArticle As String
Dim strTitleNet As String
Dim dteStarted As Date
    Dim rs As ADODB.Recordset
Dim strCode As String
    'Dim oBF As a_BookFind
    Dim strMsg As String
    Dim StartTime
    Dim x As Long
    Dim iCancelled As Integer
    Dim strShortname As String
    Dim lngProgress As Long
    Dim lngMax As Long
    iCancelled = False
Dim OpenResult As Integer
'Dim bfo As NielsenLookup
'Set bfo = New NielsenLookup
'-------------------------------
    OpenResult = oPC.OpenDBSHort
'-------------------------------
    
 '       RaiseEvent Status("Selecting records to update . . .")
        DoEvents
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
            rs.Open "SELECT * FROM vBooksWithISBNs WHERE P_EAN = '" & pEAN & "'", oPC.COShort, adOpenKeyset, adLockOptimistic
        If Not (IsISBN10(FNS(rs![P_Code])) Or IsISBN13(FNS(rs![P_EAN]))) Then
            MsgBox "Invalid code", vbInformation, "Status"
            GoTo ENDOFLOOP
        End If
        strCode = CStr(rs![P_EAN])
'        If oPC.GetProperty("NielsenUserID") > "" Then   'we are using the online service
'        If bfo Is Nothing Then
'            MsgBox "after setting bfo is nothing"
'        End If
'        xml = bfo.GetNielsenDataByISBN13(strCode, msg)
'        lngResult = ParseXmlDocument(xml, mudtProps)
 '          lngResult = ParseXmlDocument(xml, mudtProps)
'        Else
'            lngResult = oBF.FetchFromBF(pCODE)
'            If lngResult = 0 Then  'Found a record
'                lngFound = 0
'                LoadProductFromBF
'            End If
'        End If
'        If oBF.FetchFromBF(CStr(rs![P_EAN])) = 0 Then
'            If pAuthor Then rs![P_MainAuthor] = Left$(FNS(mudtProps.Author), rs.Fields("P_MainAuthor").DefinedSize)
'            rs![P_Code] = Left$(FNS(mudtProps.code), rs.Fields("P_Code").DefinedSize)
'            If pTitle Then
'                StripArticle FNS(Left$(FNS(mudtProps.Title), rs.Fields("P_Title").DefinedSize)), strArticle, strTitleNet
'                rs!P_Title = strTitleNet
'                rs!P_Article = strArticle
'            End If
'            If pSubtitle Then rs![P_SubTitle] = Left$(FNS(mudtProps.SubTitle), rs.Fields("P_SubTitle").DefinedSize)
'            If pAvailability Then rs![P_STATUS] = Left(FNS(mudtProps.Availability), 1)   ' rs.Fields("P_SubTitle").DefinedSize)
'            If pBindingcode Then rs![P_Bindingcode] = Left$(FNS(mudtProps.BindingCode), rs.Fields("P_Bindingcode").DefinedSize)
'            If pEdition Then rs![P_Edition] = Left$(FNS(mudtProps.Edition), rs.Fields("P_Edition").DefinedSize)
'            If pPublisherName Then rs![P_Publisher] = Left$(FNS(mudtProps.Publisher), rs.Fields("P_Publisher").DefinedSize)
'         '   If pSUpplierCode Then rs![P_BFSupplierCode] = Left$(FNS(mudtProps.dis), rs.fields("P_BFSupplierCode").DefinedSize)
'            strShortname = FNS(rs![P_Publisher])
'            If pUKPrice Then
'                If IsNumeric(mudtProps.UKPrice) Then
'                    rs![P_UKPrice] = CCur(mudtProps.UKPrice) * oPC.Configuration.DefaultCurrency.Divisor
'                Else
'                    rs![P_UKPrice] = Null
'                End If
'                If IsNumeric(mudtProps.USPrice) Then
'                    rs![P_USPrice] = CCur(mudtProps.USPrice) * oPC.Configuration.DefaultCurrency.Divisor
'                Else
'                    rs![P_USPrice] = Null
'                End If
'            End If
'            If pRRP Then
'                If IsNumeric(mudtProps.RRP) Then
'                    If mudtProps.RRP > 0 Then
'                        rs![P_RRP] = FNN(mudtProps.RRP * oPC.Configuration.DefaultCurrency.Divisor)
'                    Else
'                        rs![P_RRP] = FNN(mudtProps.UKPrice / IIf(oPC.Configuration.Currencies.FindBySysname("GBP") Is Nothing, 10, oPC.Configuration.Currencies.FindBySysname("GBP").Factor) * oPC.Configuration.DefaultCurrency.Divisor)
'                    End If
'                End If
'            End If
'            If pSeriesTitle Then rs![P_SeriesTitle] = Left$(FNS(mudtProps.SeriesTitle), rs.Fields("P_SeriesTitle").DefinedSize)
'            If pPublicationDate Then rs![P_Pubdate] = Left$(FNS(mudtProps.PublicationDate), rs.Fields("P_Pubdate").DefinedSize)
'            rs![P_Weight] = Left$(FNS(mudtProps.Weight), rs.Fields("P_Weight").DefinedSize)
'            If pBIC Then rs![P_BIC] = Left$(FNS(mudtProps.BFClassification), rs![P_BIC].DefinedSize)
'            rs.Update
'        Else
'            MsgBox "No record found on Nielsen", vbInformation, "Status"
'        End If
ENDOFLOOP:
        
    rs.Close
'---------------------------------------------------
    If OpenResult = 0 Then oPC.DisconnectDBShort
'---------------------------------------------------
    Exit Function

Errh::
    MsgBox Error
    Exit Function
    Resume
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProductPrev.UpdateFromBookfind(pAuthor,pTitle,pSubtitle,pAvailability,pBindingcode," & _
'        "pEdition,pSUpplierCode,pPublisherName,pSeriesTitle,pPublicationDate,pUKPrice,pRRP,pBIC,pBookStatus," & _
'        "pSTAFFID,pEAN)", Array(pAuthor, pTitle, pSubtitle, pAvailability, pBindingcode, pEdition, _
'         pSUpplierCode, pPublisherName, pSeriesTitle, pPublicationDate, pUKPrice, pRRP, pBIC, pBookStatus, _
'         pSTAFFID, pEAN)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.UpdateFromBookfind(pAuthor,pTitle,pSubtitle,pAvailability,pBindingcode," & _
        "pEdition,pSUpplierCode,pPublisherName,pSeriesTitle,pPublicationDate,pUKPrice,pRRP,pBIC,pBookStatus," & _
        "pSTAFFID,pEAN)", Array(pAuthor, pTitle, pSubtitle, pAvailability, pBindingcode, pEdition, _
         pSUpplierCode, pPublisherName, pSeriesTitle, pPublicationDate, pUKPrice, pRRP, pBIC, pBookStatus, _
         pSTAFFID, pEAN)
End Function

Private Sub cmdDropAppros_Click()
    On Error GoTo errHandler
    APPGRID.ZOrder 0
    If lngAPPGRIDHeight = 0 Then
        lngAPPGRIDHeight = APPGRID.Height
        APPGRID.Height = APPGRID.Height * 2
    Else
        APPGRID.Height = lngAPPGRIDHeight
        lngAPPGRIDHeight = 0
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdDropAppros_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDropCO_Click()
    On Error GoTo errHandler
    COGrid.ZOrder 0
    If lngCOGridHeight = 0 Then
        lngCOGridHeight = COGrid.Height
        COGrid.Height = COGrid.Height * 1.35
    Else
        COGrid.Height = lngCOGridHeight
        lngCOGridHeight = 0
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdDropCO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDropMovements_Click()
    On Error GoTo errHandler
    MMGRID.ZOrder 0
    If lngMMGRIDHeight = 0 Then
        lngMMGRIDHeight = MMGRID.Height
        MMGRID.Height = MMGRID.Height * 1.35
    Else
        MMGRID.Height = lngMMGRIDHeight
        lngMMGRIDHeight = 0
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdDropMovements_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDropPO_Click()
    On Error GoTo errHandler
    POGrid.ZOrder 0
    If lngPOGridHeight = 0 Then
        lngPOGridHeight = POGrid.Height
        POGrid.Height = POGrid.Height * 2
    Else
        POGrid.Height = lngPOGridHeight
        lngPOGridHeight = 0
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdDropPO_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdDropStock_Click()
    On Error GoTo errHandler
    StGrid.ZOrder 0
    If lngSTGridHeight = 0 Then
        lngSTGridHeight = StGrid.Height
        StGrid.Height = StGrid.Height * 3
    Else
        StGrid.Height = lngSTGridHeight
        lngSTGridHeight = 0
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrev.cmdDropStock_Click", , EA_NORERAISE
    HandleError
End Sub

