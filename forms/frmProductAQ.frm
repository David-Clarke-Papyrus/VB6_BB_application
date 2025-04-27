VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmProductAQ 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Stock"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11565
   ControlBox      =   0   'False
   Icon            =   "frmProductAQ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleMode       =   0  'User
   ScaleWidth      =   15255.96
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
      Left            =   10380
      Picture         =   "frmProductAQ.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   5820
      Width           =   1000
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
      Left            =   9360
      Picture         =   "frmProductAQ.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   5820
      Width           =   1000
   End
   Begin VB.TextBox txtFlag 
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
      Left            =   6420
      TabIndex        =   9
      Top             =   1860
      Width           =   4905
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
      Height          =   1500
      Left            =   9060
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   285
      Width           =   2250
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   6105
      Width           =   4350
   End
   Begin VB.CommandButton cmdNewCode 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&New code"
      Height          =   420
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Visible         =   0   'False
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3480
      Left            =   120
      TabIndex        =   11
      Top             =   2205
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   6138
      _Version        =   393216
      Style           =   1
      TabsPerRow      =   5
      TabHeight       =   468
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
      TabCaption(0)   =   "&1. Copies"
      TabPicture(0)   =   "frmProductAQ.frx":0A1E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdRemove"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdAddCopy"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSaveLayout"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProductAQ.frx":0A3A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdRefresh"
      Tab(1).Control(1)=   "txtSP"
      Tab(1).Control(2)=   "txtRRP"
      Tab(1).Control(3)=   "txtEAN"
      Tab(1).Control(4)=   "txtVAT"
      Tab(1).Control(5)=   "cmdSetDefault"
      Tab(1).Control(6)=   "chkServiceItem"
      Tab(1).Control(7)=   "chkObsolete"
      Tab(1).Control(8)=   "cboCatHead"
      Tab(1).Control(9)=   "cboSection"
      Tab(1).Control(10)=   "txtBinding"
      Tab(1).Control(11)=   "txtBIC"
      Tab(1).Control(12)=   "Label18"
      Tab(1).Control(13)=   "Label16"
      Tab(1).Control(14)=   "Label11"
      Tab(1).Control(15)=   "Label26"
      Tab(1).Control(16)=   "Label10"
      Tab(1).Control(17)=   "Label17"
      Tab(1).Control(18)=   "Label20"
      Tab(1).Control(19)=   "Label25"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "&3. Wants"
      TabPicture(2)   =   "frmProductAQ.frx":0A56
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdNewWant"
      Tab(2).Control(1)=   "lvwWants"
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   -66360
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdSaveLayout 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Save layout"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1155
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3120
         Width           =   1665
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
         Left            =   -67830
         TabIndex        =   44
         Top             =   2250
         Visible         =   0   'False
         Width           =   1380
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
         Left            =   -67845
         TabIndex        =   43
         Top             =   1875
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txtEAN 
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
         Left            =   -68175
         TabIndex        =   41
         Top             =   1380
         Width           =   1725
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
         Left            =   -66330
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   525
         Width           =   495
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
         Height          =   300
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3090
         Width           =   420
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
         Height          =   300
         Left            =   570
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3090
         Width           =   405
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   2670
         Left            =   120
         OleObjectBlob   =   "frmProductAQ.frx":0A72
         TabIndex        =   37
         Top             =   405
         Width           =   10965
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
         Left            =   -72585
         TabIndex        =   31
         Top             =   1605
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
         Left            =   -71130
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1575
         Width           =   1755
      End
      Begin VB.CheckBox chkServiceItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Non-stock"
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
         Left            =   -73605
         TabIndex        =   29
         Top             =   1995
         Width           =   1245
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
         Height          =   480
         Left            =   -73605
         TabIndex        =   28
         Top             =   2370
         Width           =   1245
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
         Left            =   -72585
         TabIndex        =   27
         Text            =   "cboCatHead"
         Top             =   990
         Width           =   6135
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
         Left            =   -70200
         TabIndex        =   26
         Top             =   2925
         Visible         =   0   'False
         Width           =   3420
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
         Left            =   -65160
         TabIndex        =   25
         Top             =   1710
         Visible         =   0   'False
         Width           =   1395
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
         Left            =   -65145
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   1380
      End
      Begin MSComctlLib.ListView lvwWants 
         Height          =   2430
         Left            =   -74760
         TabIndex        =   39
         Top             =   540
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   4286
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
            Name            =   "Arial Narrow"
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
         Left            =   -68640
         TabIndex        =   46
         Top             =   2265
         Visible         =   0   'False
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
         Left            =   -68640
         TabIndex        =   45
         Top             =   1890
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
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
         Left            =   -69420
         TabIndex        =   42
         Top             =   1410
         Width           =   1080
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
         Left            =   -71445
         TabIndex        =   36
         Top             =   2940
         Visible         =   0   'False
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
         Left            =   -73815
         TabIndex        =   35
         Top             =   1635
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
         Left            =   -74490
         TabIndex        =   34
         Top             =   1035
         Width           =   1755
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
         Left            =   -66390
         TabIndex        =   33
         Top             =   1725
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "B.I.C"
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
         Left            =   -66390
         TabIndex        =   32
         Top             =   1350
         Visible         =   0   'False
         Width           =   1080
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
      Left            =   6420
      TabIndex        =   5
      Top             =   150
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
      Left            =   6420
      TabIndex        =   6
      Top             =   585
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
      Left            =   6420
      TabIndex        =   8
      Top             =   1455
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
      Left            =   6420
      TabIndex        =   7
      Top             =   1020
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
      TabIndex        =   4
      Top             =   1545
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
      Left            =   735
      TabIndex        =   2
      Top             =   525
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
      Left            =   750
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   930
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
      Top             =   90
      Width           =   1680
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Height          =   255
      Left            =   5400
      TabIndex        =   40
      Top             =   1920
      Width           =   945
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
      Left            =   9045
      TabIndex        =   23
      Top             =   60
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
      Left            =   4785
      TabIndex        =   21
      Top             =   165
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
      Left            =   4890
      TabIndex        =   20
      Top             =   615
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
      Left            =   5715
      TabIndex        =   19
      Top             =   1515
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
      Left            =   5385
      TabIndex        =   18
      Top             =   1065
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
      Left            =   30
      TabIndex        =   17
      Top             =   570
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
      Left            =   30
      TabIndex        =   16
      Top             =   1560
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
      Left            =   180
      TabIndex        =   15
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   30
      TabIndex        =   14
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmProductAQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents oProd As a_Product
Attribute oProd.VB_VarHelpID = -1
Dim flgLoading As Boolean
'Dim tlCatHead As z_TextList
'Private tlSections As z_TextList
'Private tlProductTypes As z_TextList
Dim mCancel As Boolean
Dim XA As XArrayDB
Dim frmPrevious As Form

Sub component(pProduct As a_Product, Optional pPrevForm As Form)
    On Error GoTo errHandler
    Set frmPrevious = pPrevForm
    Set oProd = pProduct
    oProd.BeginEdit
    oProd.SetBook
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.component(pProduct,pPrevForm)", Array(pProduct, pPrevForm)
End Sub


Private Sub cboSection_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetCategoryID oPC.Configuration.Sections.Key(cboSection)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cboSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboCatHead_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetCatalogueheadingID oPC.Configuration.CatalogueHeadings.Key(cboCatHead)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cboCatHead_Click", , EA_NORERAISE
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
        Set tmpCopy = oProd.Copies(XA.Value(Grid1.Bookmark, 9))
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdAddCopy_Click", , EA_NORERAISE
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
    ErrorIn "frmProductAQ.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdNewWant_Click()
    On Error GoTo errHandler
Dim frm As frmWant
'Dim oWant As a_Want
   ' Set oWant = oProd.Wants.Add
    Set frm = New frmWant
    frm.component oProd
    frm.Show vbModal
    oProd.Wants.Load oProd.PID, enWant
    LoadWants
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdNewWant_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo errHandler
    oPC.Configuration.ReloadCatHeads
    LoadCombo cboCatHead, oPC.Configuration.CatalogueHeadings
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdRefresh_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdRemove_Click()
    On Error GoTo errHandler
Dim oCopy As a_Copy
    If Grid1.row = -1 Then Exit Sub
    Set oCopy = oProd.Copies.FindBySerial(XA(Grid1.Bookmark, 1))
    oCopy.BeginEdit
    oCopy.Delete
    oCopy.ApplyEdit
    LoadCopies
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdRemove_Click", , EA_NORERAISE
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.UPdateFromXArray"
End Sub

Private Sub cmdSaveLayout_Click()
    On Error GoTo errHandler
Dim i As Integer
    For i = 1 To Grid1.Columns.Count
        SaveSetting "PBKS", Me.Name, CStr(i), Grid1.Columns(i - 1).Width
    Next
    SaveSetting "PBKS", Me.Name, "Rowheight", Grid1.RowHeight

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdSaveLayout_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdRemoveWant_Click()
'Dim oWant As a_Want
'Dim lngCOLID As Long
'
'    If lvwWants.SelectedItem Is Nothing Then Exit Sub
'    lngCOLID = oProd.Wants(lvwWants.SelectedItem.Key).COLID
'    oWant.BeginEdit
'    oWant.Delete
'    oWant.ApplyEdit
'    LoadWants
'End Sub

Private Sub cmdSetDefault_Click()
    On Error GoTo errHandler
    Me.txtVAT = oPC.Configuration.VATRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdSetDefault_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oProd.IsEditing Then oProd.CancelEdit
'    If frmProductPrevAQ.WindowState > 0 Then
     '   frmProductPrevAQ.RefreshForm
 '   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
Dim i As Integer
Dim oC As a_Copy
    i = ColIndex + 1
    Set oC = oProd.Copies(XA(Grid1.Bookmark, 9))
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, OldValue, _
         Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
'MsgBox "Selected row is : " & Grid1.Row + 1
Dim frm As frmCopy
Dim tmpCopy As a_Copy
Dim oCopy As a_Copy
    Set oCopy = oProd.Copies(XA(Grid1.Bookmark, 9))
    Set frm = New frmCopy
    If Grid1.Bookmark > 0 Then
        Set tmpCopy = oProd.Copies(Grid1.Bookmark)
    Else
        Set tmpCopy = Nothing
    End If
    frm.component oCopy, tmpCopy
    frm.Show vbModal
    LoadCopies
    Grid1.ReBind

    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProductAQ: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmProductAQ: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 7) > "" Then
        RowStyle.BackColor = &HDCDBF2
    End If
    If XA(Bookmark, 10) = True Then
        RowStyle.BackColor = vbRed
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub lvwWants_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.lvwWants_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwWants_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.lvwWants_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwWants_DblClick()
    On Error GoTo errHandler
Dim frm As frmCOPreview
    If lvwWants.SelectedItem Is Nothing Then Exit Sub
    Set frm = New frmCOPreview
    frm.component oProd.Wants.Item(lvwWants.SelectedItem.Key).TRID, False
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.lvwWants_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub oProd_Valid(strMsg As String)
    On Error GoTo errHandler
    Me.txtErrors = strMsg
    Me.cmdOK.Enabled = (strMsg = "")
    Me.cmdAddCopy.Enabled = (strMsg = "")
    Me.cmdRemove.Enabled = (strMsg = "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.oProd_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oProd.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdNewCode_Click()
    On Error GoTo errHandler
    Me.txtCode = "#"
    oProd.SetCode "#"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdNewCode_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim frmPreview As Form
    oProd.ApplyEdit lngResult
    If lngResult = 99 Then
        MsgBox "Invalid values - check that the code is has not been already used", , "Save failed"
    Else
        If frmPrevious Is Nothing Then
        '    If oPC.Configuration.AntiquarianYN Then
                Set frmPreview = New frmProductPrevAQ
        '    Else
        '        Set frmPreview = New frmProductPrev
        '    End If
        Else
            Set frmPreview = frmPrevious
        End If
        frmPreview.component oProd
        frmPreview.RefreshForm
        frmPreview.Show
        
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11700
        Height = 7050
    End If
    LoadControls
    Me.cmdNewCode.Enabled = oProd.IsNew
    Me.SSTab1.Tab = 0
    LoadWants
    oProd.GetStatus
    If oProd.IsNew Then
        Me.txtCode = "#"
        mSetfocus Me.txtAuthor
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtCode = oProd.code
    txtEAN = oProd.EAN
    txtTitle = oProd.Title
    txtSubtitle = oProd.SubTitle
    txtAuthor = oProd.Author
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    txtPubPlace = oProd.PublicationPlace
    Me.txtPubDate = oProd.PublicationDate
    Me.txtBinding = oProd.BindingCode
    Me.txtFlag = oProd.FlagText
    txtBIC = oProd.BIC
    txtRRP = oProd.RRPF
    txtSP = oProd.SPF
    Me.txtNote = oProd.Note
    Me.txtVAT = oProd.VATRateF
    LoadCombo cboCatHead, oPC.Configuration.CatalogueHeadings
    LoadCombo cboSection, oPC.Configuration.ProductTypes
    cboSection = oPC.Configuration.ProductTypes.Item(oProd.CategoryID)
    If oProd.CatalogueheadingID > 0 Then cboCatHead = oPC.Configuration.CatalogueHeadings.Item(oProd.CatalogueheadingID)
    Me.chkServiceItem = IIf(oProd.IsServiceItem, 1, 0)
    Me.chkObsolete = IIf(oProd.Obsolete, 1, 0)
    LoadCopies
    LoadWants
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.LoadControls"
End Sub
Private Sub LoadCopies()
    On Error GoTo errHandler
Dim objItm As ListItem
Dim lngIndex As Long
Dim i, j As Integer
Dim tmp As String
Dim strCatalogues As String

    For i = 1 To Grid1.Columns.Count
        Grid1.Columns(i - 1).Width = GetSetting("PBKS", Me.Name, CStr(i), Grid1.Columns(i - 1).Width)
    Next
    Grid1.RowHeight = GetSetting("PBKS", Me.Name, "Rowheight", 270)
    
    Set XA = New XArrayDB
    XA.Clear
    XA.ReDim 1, oProd.Copies.Count, 1, 11
    For lngIndex = 1 To oProd.Copies.Count
        strCatalogues = ""
        For j = 1 To oProd.Copies(lngIndex).CatalogueEntries.Count
            strCatalogues = strCatalogues & oProd.Copies(lngIndex).CatalogueEntries(j).Serial
            If j < oProd.Copies(lngIndex).CatalogueEntries.Count Then strCatalogues = strCatalogues & ", "
        Next j
        XA.Value(lngIndex, 1) = oProd.Copies(lngIndex).Serial
        XA.Value(lngIndex, 2) = oProd.Copies(lngIndex).Description
        XA.Value(lngIndex, 3) = oProd.Copies(lngIndex).Comment
        XA.Value(lngIndex, 4) = oProd.Copies(lngIndex).CatalogueEntries_Concat
        XA.Value(lngIndex, 5) = oProd.Copies(lngIndex).SoldTo
        XA.Value(lngIndex, 6) = oProd.Copies(lngIndex).PurchaseDateF
        XA.Value(lngIndex, 7) = oProd.Copies(lngIndex).SoldDateF
        XA.Value(lngIndex, 8) = oProd.Copies(lngIndex).PriceF
        XA.Value(lngIndex, 9) = oProd.Copies(lngIndex).Key
        XA.Value(lngIndex, 10) = oProd.Copies(lngIndex).IsDeleted
        XA.Value(lngIndex, 11) = IIf((oProd.Copies(lngIndex).Serial = 0), 1, 0)
    Next
    XA.QuickSort 1, oProd.Copies.Count, 11, XORDER_DESCEND, XTYPE_INTEGER, 1, XORDER_DESCEND, XTYPE_INTEGER
    Grid1.Array = XA
 '   Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.LoadCopies"
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.LoadWants"
End Sub

Private Sub lvwCopies_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.lvwCopies_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtBIC_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtBIC = oProd.BIC
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtBIC_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBIC_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtBIC_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBIC_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductAQ.txtBIC_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtBinding_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtBinding = oProd.BindingCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtBinding_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtBinding_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtBinding_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtBinding_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductAQ.txtBinding_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oProd.SetCode txtCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtEAN_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtEAN = oProd.EAN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtEAN_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEAN_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtEAN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtEAN_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductAQ.txtEAN_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtFlag_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtFlag = oProd.FlagText
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtFlag_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtFlag_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtFlag_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtFlag_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetFlagtext(txtFlag)
    If Err Then
      Beep
      intPos = txtFlag.SelStart
      txtFlag = oProd.FlagText
      txtFlag.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtFlag_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtRRP_GotFocus()
    On Error GoTo errHandler
    txtRRP = oProd.RRP
    AutoSelect txtRRP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtRRP_GotFocus", , EA_NORERAISE
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
    ErrorIn "frmProductAQ.txtRRP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtSP_GotFocus()
    On Error GoTo errHandler
    txtSP = oProd.SP
    AutoSelect txtSP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtSP_GotFocus", , EA_NORERAISE
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
    ErrorIn "frmProductAQ.txtSP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtSubtitle_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtSubtitle = oProd.SubTitle
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtSubtitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSubtitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtSubtitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtSubtitle_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetSubTitle(txtSubtitle)
    If Err Then
      Beep
      intPos = txtSubtitle.SelStart
      txtSubtitle = oProd.SubTitle
      txtSubtitle.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtSubtitle_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtTitle = oProd.Title
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Change()
    On Error GoTo errHandler
Dim intPos As Integer
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
    ErrorIn "frmProductAQ.txtTitle_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAuthor_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtAuthor = oProd.Author
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtAuthor_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtAuthor_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtAuthor_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtAuthor_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductAQ.txtAuthor_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPublisher = oProd.Publisher
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtPublisher_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtPublisher_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductAQ.txtPublisher_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubDate_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPubDate = oProd.PublicationDate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtPubDate_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubDate_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtPubDate_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubDate_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductAQ.txtPubDate_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubPlace_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPubPlace = oProd.PublicationPlace
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtPubPlace_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubPlace_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtPubPlace_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPubPlace_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductAQ.txtPubPlace_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtEdition = oProd.Edition
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtEdition_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtEdition_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductAQ.txtEdition_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtNote = oProd.Note
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtNote_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtNote_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtNote_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
    mCancel = Not oProd.SetNote(txtNote)
    If Err Then
      Beep
      intPos = txtNote.SelStart
      txtNote = oProd.Note
      txtNote.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtVAT_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtVAT
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtVAT_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtVAT_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtVAT = oProd.VATRateToUse
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtVAT_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtVAT_Validate(Cancel As Boolean)
    On Error GoTo errHandler
   If flgLoading Then Exit Sub
   Cancel = Not Not oProd.SetVAT(txtVAT)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductAQ.txtVAT_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

