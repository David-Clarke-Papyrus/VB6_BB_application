VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmProductPrevAQ 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Stock"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11490
   ControlBox      =   0   'False
   Icon            =   "frmProductPrevAQ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   11490
   Begin VB.CommandButton cmdCopy 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Copy to new title record"
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
      Left            =   5835
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   5850
      Width           =   2280
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
      Height          =   330
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   1785
      Width           =   4980
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Find By ISBN"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   735
      TabIndex        =   31
      Top             =   0
      Width           =   3255
      Begin VB.CommandButton cmdsearchisbn 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   270
         Width           =   1995
      End
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
      Left            =   9330
      Picture         =   "frmProductPrevAQ.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5835
      Width           =   1000
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
      Left            =   10350
      Picture         =   "frmProductPrevAQ.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5835
      Width           =   1000
   End
   Begin VB.TextBox txtNote 
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
      Height          =   750
      Left            =   9045
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   930
      Width           =   2355
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   585
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5790
      Width           =   4350
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3540
      Left            =   135
      TabIndex        =   14
      Top             =   2220
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   6244
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "&1. Copies"
      TabPicture(0)   =   "frmProductPrevAQ.frx":073F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdSaveLayout"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&2. Details"
      TabPicture(1)   =   "frmProductPrevAQ.frx":075B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRRP"
      Tab(1).Control(1)=   "txtSP"
      Tab(1).Control(2)=   "txtEAN"
      Tab(1).Control(3)=   "txtCategoryHeading"
      Tab(1).Control(4)=   "txtCategory"
      Tab(1).Control(5)=   "txtVAT"
      Tab(1).Control(6)=   "txtBinding"
      Tab(1).Control(7)=   "txtBIC"
      Tab(1).Control(8)=   "Label16"
      Tab(1).Control(9)=   "Label18"
      Tab(1).Control(10)=   "Label11"
      Tab(1).Control(11)=   "lblObsolete"
      Tab(1).Control(12)=   "lblServiceItem"
      Tab(1).Control(13)=   "Label26"
      Tab(1).Control(14)=   "Label10"
      Tab(1).Control(15)=   "Label17"
      Tab(1).Control(16)=   "Label20"
      Tab(1).Control(17)=   "Label25"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "&3. Wants"
      TabPicture(2)   =   "frmProductPrevAQ.frx":0777
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lvwWants"
      Tab(2).ControlCount=   1
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
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3210
         Width           =   1665
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
         Left            =   -67770
         TabIndex        =   44
         Top             =   1845
         Visible         =   0   'False
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
         Left            =   -67770
         TabIndex        =   43
         Top             =   2220
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txtEAN 
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
         Left            =   -68445
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1035
         Width           =   1890
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
         Left            =   -72765
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   6210
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
         Left            =   -65820
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1665
         Visible         =   0   'False
         Width           =   2115
      End
      Begin TrueOleDBGrid60.TDBGrid Grid1 
         Height          =   2790
         Left            =   120
         OleObjectBlob   =   "frmProductPrevAQ.frx":0793
         TabIndex        =   28
         Top             =   390
         Width           =   10965
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
         Left            =   -72750
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1065
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
         Left            =   -65820
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2850
         Visible         =   0   'False
         Width           =   1395
      End
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
         Left            =   -65820
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2460
         Visible         =   0   'False
         Width           =   1395
      End
      Begin MSComctlLib.ListView lvwWants 
         Height          =   2460
         Left            =   -74760
         TabIndex        =   34
         Top             =   510
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   4339
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483635
         BackColor       =   14155263
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
         Left            =   -68580
         TabIndex        =   46
         Top             =   1860
         Visible         =   0   'False
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
         Left            =   -68580
         TabIndex        =   45
         Top             =   2235
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
         Left            =   -69675
         TabIndex        =   42
         Top             =   1065
         Width           =   1080
      End
      Begin VB.Label lblObsolete 
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -72600
         TabIndex        =   38
         Top             =   2955
         Width           =   1380
      End
      Begin VB.Label lblServiceItem 
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   -72600
         TabIndex        =   37
         Top             =   2625
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
         Left            =   -67050
         TabIndex        =   27
         Top             =   1680
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
         Left            =   -73980
         TabIndex        =   26
         Top             =   1095
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
         Left            =   -74670
         TabIndex        =   25
         Top             =   630
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
         Left            =   -67050
         TabIndex        =   24
         Top             =   2865
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
         Left            =   -67050
         TabIndex        =   23
         Top             =   2490
         Visible         =   0   'False
         Width           =   1080
      End
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
      Left            =   6405
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   60
      Width           =   2520
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
      Left            =   6405
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   495
      Width           =   2520
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
      Left            =   6405
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1365
      Width           =   2520
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
      Left            =   6405
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   930
      Width           =   2520
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
      Height          =   510
      Left            =   735
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1665
      Width           =   3900
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
      Height          =   330
      Left            =   735
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   735
      Width           =   3900
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
      Height          =   510
      Left            =   735
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1110
      Width           =   3900
   End
   Begin VB.TextBox txtCode 
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
      Height          =   330
      Left            =   9690
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   45
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
      Top             =   1845
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
      Height          =   195
      Left            =   9060
      TabIndex        =   19
      Top             =   675
      Width           =   390
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
      Left            =   4770
      TabIndex        =   16
      Top             =   75
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
      Left            =   4875
      TabIndex        =   15
      Top             =   525
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
      Left            =   5700
      TabIndex        =   13
      Top             =   1425
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
      Left            =   5370
      TabIndex        =   12
      Top             =   975
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
      TabIndex        =   11
      Top             =   720
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
      TabIndex        =   10
      Top             =   1665
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
      TabIndex        =   9
      Top             =   1170
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
      Left            =   8970
      TabIndex        =   8
      Top             =   75
      Width           =   660
   End
End
Attribute VB_Name = "frmProductPrevAQ"
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
Dim XA As XArrayDB

Sub component(pProduct As a_Product)
    On Error GoTo errHandler
    Set oProd = pProduct
    Set tlCatHead = New z_TextList
    tlCatHead.Load ltCatalogueHeadings
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.component(pProduct)", pProduct
End Sub


Private Sub cboCatHead_Click()
    On Error GoTo errHandler
'    oProd.setCatalogueheadingID tlCatHead.Key(cboCatHead)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.cboCatHead_Click", , EA_NORERAISE
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
    ErrorIn "frmProductPrevAQ.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdCopy_Click()
    On Error GoTo errHandler
Dim oNewProd As a_Product
Dim frmA As frmProductPrevAQ

    If MsgBox("You wish create a new title record with this title's data?", vbQuestion + vbOKCancel, "Confirm") = vbCancel Then
        Exit Sub
    End If
    Set oNewProd = New a_Product
    oNewProd.BeginEdit
    
    oNewProd.SetAuthor oProd.Author
    oNewProd.SetBindingCode oProd.BindingCode
    oNewProd.SetCatalogueheadingID oProd.CatalogueheadingID
    oNewProd.SetCategoryID oProd.CategoryID
    oNewProd.SetComment oProd.Comment
    oNewProd.SetCost oProd.Cost
    oNewProd.SetDescription oProd.Description
 '   oNewProd.SetEAN oProd.EAN
    oNewProd.SetEdition oProd.Edition
    oNewProd.SetFlagtext oProd.FlagText
    oNewProd.ForeignOrderedCURRID = oProd.ForeignOrderedCURRID
    oNewProd.SetForeignOrderedPrice oProd.ForeignOrderedPrice
    oNewProd.LoyaltyRate = oProd.LoyaltyRate
    oNewProd.SetNote oProd.Note
    oNewProd.Obsolete = oProd.Obsolete
    oNewProd.SetProductType oProd.ProductType
    oNewProd.SetProductTypeID oProd.ProductTypeID
    oNewProd.SetPublicationDate oProd.PublicationDate
    oNewProd.SetPublicationPlace oProd.PublicationPlace
    oNewProd.SetPublisher oProd.Publisher
    oNewProd.SetRRP oProd.RRP
 '   oNewProd.SetSection oProd.Section
    oNewProd.Seesafe = oProd.Seesafe
    oNewProd.SetSeriesTitle oProd.SeriesTitle
    oNewProd.SetBIC oProd.BIC
    oNewProd.SetComment oProd.Comment
    oNewProd.SetSP oProd.SP
  '  oNewProd.SetStatus oProd.Status
    oNewProd.SetSubTitle oProd.SubTitle
    oNewProd.SetSummary oProd.Summary
    oNewProd.SetTitle oProd.Title
    oNewProd.SetUKPrice oProd.UKPrice
    oNewProd.SetUSPrice oProd.USPrice
    oNewProd.SetVAT oProd.VATRate
    oNewProd.SpecialVat = oProd.SpecialVat
    oNewProd.SupplierID = oProd.SupplierID
    
    oNewProd.ApplyEdit
    Set frmA = New frmProductPrevAQ
    frmA.component oNewProd
    frmA.Show
    Unload Me
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.cmdCopy_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo errHandler
Dim frm As frmProductAQ
    Set frm = New frmProductAQ
    frm.component oProd, Me
    frm.Show
    
  '  Unload Me
  '  Set oProd = Nothing
  '  Set frm = Nothing
    Exit Sub
Errh:
    MsgBox Error

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.cmdEdit_Click", , EA_NORERAISE
    HandleError
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
    ErrorIn "frmProductPrevAQ.cmdSaveLayout_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdsearchisbn_Click()
    On Error GoTo errHandler
    Set oProd = Nothing
    Set oProd = New a_Product
    With oProd
    .Load "", 0, txtisbnsearch
       
    txtAuthor = .Author
    txtCode = .code
    txtSubtitle = .SubTitle
    txtTitle = .Title
    txtPublisher = .Publisher
        
    End With
    LoadCopies
    Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.cmdsearchisbn_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Command1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Command1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Activate()
    On Error GoTo errHandler
    RefreshForm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Form_Activate", , EA_NORERAISE
    HandleError
End Sub
Private Sub Form_DblClick()
    On Error GoTo errHandler

    If Not IsNull(oProd) Then
        On Error Resume Next
        Clipboard.Clear
        Clipboard.SetText oProd.ProductDetails
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Form_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub Grid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Grid1_BeforeColEdit(ColIndex,KeyAscii,Cancel)", Array(ColIndex, _
         KeyAscii, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    On Error GoTo errHandler
         '   OldValue = Grid1.Text
            Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Grid1_BeforeColUpdate(ColIndex,OldValue,Cancel)", Array(ColIndex, _
         OldValue, Cancel), EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_DblClick()
    On Error GoTo errHandler
'MsgBox "Selected row is : " & Grid1.Row + 1
Dim frm As frmCopyPreview
Dim oCopy As a_Copy
    If IsNull(Grid1.Bookmark) Then Exit Sub
    Set oCopy = oProd.Copies(val(XA(Grid1.Bookmark, 9)))
    Set frm = New frmCopyPreview

    frm.component oCopy, oProd
    frm.Show ' vbModal

    Exit Sub
errHandler:
    ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmProductPrevAQ: Grid1_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmProductPrevAQ: Grid1_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
     If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Grid1_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub Grid1_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If XA(Bookmark, 7) > "" Then
        RowStyle.BackColor = &HDCDBF2
    End If
    If XA(Bookmark, 7) = True Then
        RowStyle.BackColor = vbRed
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Grid1_FetchRowStyle(Split,Bookmark,RowStyle)", Array(Split, Bookmark, _
         RowStyle), EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 11700
        Height = 7100
    End If
    LoadControls
    Me.SSTab1.Tab = 0
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Public Sub RefreshForm()
    On Error GoTo errHandler
    LoadControls
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.RefreshForm"
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
    Me.txtRRP = oProd.RRPF
    Me.txtSP = oProd.SPF
    Me.txtPubDate = oProd.PublicationDate
    Me.txtBinding = oProd.BindingCode
    Me.txtCategory = oPC.Configuration.Sections.Item(oProd.CategoryID)
    Me.txtCategoryHeading = tlCatHead.Item(oProd.CatalogueheadingID)
    Me.txtNote = oProd.Note
    Me.txtVAT = oProd.VATRateF
    txtFlagText = oProd.FlagText
    Me.txtBinding = oProd.BindingCode
    txtBIC = oProd.BIC
    Me.lblServiceITem = IIf(oProd.IsServiceItem, "non-stock", "")
    Me.lblObsolete = IIf(oProd.Obsolete, "obsolete", "")
    LoadCopies
    LoadWants
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.LoadControls"
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
        XA.Value(lngIndex, 11) = IIf(oProd.Copies(lngIndex).SoldDateF > "", 1, 0)
    Next
    XA.QuickSort 1, oProd.Copies.Count, 11, XORDER_ASCEND, XTYPE_INTEGER, 1, XORDER_DESCEND, XTYPE_INTEGER
    Grid1.Array = XA
    Grid1.ReBind
 '   Grid1.ReBind
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.LoadCopies"
End Sub

Private Sub lvwCopies_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.lvwCopies_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
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
    ErrorIn "frmProductPrevAQ.LoadWants"
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
    If Shift = 2 And KeyCode = vbKeyA Then
        Grid1.EditActive = True
        Grid1.SelStart = 0
        Grid1.SelLength = Len(Grid1.text) - 1
        Grid1.Refresh
        Grid1.EditActive = False
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.Grid1_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub lvwWants_AfterLabelEdit(Cancel As Integer, NewString As String)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.lvwWants_AfterLabelEdit(Cancel,NewString)", Array(Cancel, NewString), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub lvwWants_BeforeLabelEdit(Cancel As Integer)
    On Error GoTo errHandler
Cancel = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.lvwWants_BeforeLabelEdit(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub lvwWants_DblClick()
    On Error GoTo errHandler
Dim frm As frmCOPreview
    If lvwWants.SelectedItem.Index < 1 Then Exit Sub
    Set frm = New frmCOPreview
    frm.component oProd.Wants.Item(lvwWants.SelectedItem.Key).TRID, False
    frm.Show
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductPrevAQ.lvwWants_DblClick", , EA_NORERAISE
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
    ErrorIn "frmProductPrevAQ.ExportInCatalogueFormat"
End Sub
