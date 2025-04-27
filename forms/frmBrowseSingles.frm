VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Begin VB.Form frmBrowseSingles 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Browse products"
   ClientHeight    =   6825
   ClientLeft      =   240
   ClientTop       =   1020
   ClientWidth     =   15900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrowseSingles.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   15900
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Export"
      Height          =   615
      Left            =   4605
      Picture         =   "frmBrowseSingles.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5940
      Width           =   1000
   End
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   5640
      Left            =   4590
      OleObjectBlob   =   "frmBrowseSingles.frx":0914
      TabIndex        =   9
      Top             =   255
      Width           =   11130
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   11745
      Picture         =   "frmBrowseSingles.frx":4FAF
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5985
      Width           =   1000
   End
   Begin VB.CommandButton cmdSaveLayout 
      BackColor       =   &H00C4BCA4&
      Caption         =   "save layout"
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
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6525
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Height          =   6495
      Left            =   75
      TabIndex        =   3
      Top             =   135
      Width           =   4425
      Begin VB.TextBox txtLengthMargin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1965
         TabIndex        =   42
         Top             =   4905
         Width           =   390
      End
      Begin VB.TextBox txtWidthMargin 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1965
         TabIndex        =   41
         Top             =   5220
         Width           =   390
      End
      Begin VB.TextBox txtDescriptionOrCode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1380
         TabIndex        =   39
         Top             =   270
         Width           =   2565
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2730
         TabIndex        =   37
         Top             =   4905
         Width           =   1200
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   885
         TabIndex        =   33
         Top             =   5220
         Width           =   945
      End
      Begin VB.TextBox txtLength 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   885
         TabIndex        =   32
         Top             =   4905
         Width           =   945
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   7
         ItemData        =   "frmBrowseSingles.frx":5339
         Left            =   1365
         List            =   "frmBrowseSingles.frx":533B
         TabIndex        =   23
         Top             =   4245
         Width           =   2640
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         ItemData        =   "frmBrowseSingles.frx":533D
         Left            =   1365
         List            =   "frmBrowseSingles.frx":533F
         TabIndex        =   22
         Top             =   3885
         Width           =   2640
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         ItemData        =   "frmBrowseSingles.frx":5341
         Left            =   1365
         List            =   "frmBrowseSingles.frx":5343
         TabIndex        =   21
         Top             =   3540
         Width           =   2640
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         ItemData        =   "frmBrowseSingles.frx":5345
         Left            =   1365
         List            =   "frmBrowseSingles.frx":5347
         TabIndex        =   20
         Top             =   3180
         Width           =   2640
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         ItemData        =   "frmBrowseSingles.frx":5349
         Left            =   1365
         List            =   "frmBrowseSingles.frx":534B
         TabIndex        =   19
         Top             =   2820
         Width           =   2640
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         ItemData        =   "frmBrowseSingles.frx":534D
         Left            =   1365
         List            =   "frmBrowseSingles.frx":534F
         TabIndex        =   18
         Top             =   2460
         Width           =   2640
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         ItemData        =   "frmBrowseSingles.frx":5351
         Left            =   1365
         List            =   "frmBrowseSingles.frx":5353
         TabIndex        =   17
         Top             =   2115
         Width           =   2640
      End
      Begin VB.ComboBox cboSearch 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         ItemData        =   "frmBrowseSingles.frx":5355
         Left            =   1365
         List            =   "frmBrowseSingles.frx":5357
         TabIndex        =   15
         Top             =   1755
         Width           =   2640
      End
      Begin VB.CommandButton cmdClearSection 
         BackColor       =   &H00D3C9C0&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1110
         Width           =   255
      End
      Begin VB.CommandButton cmdClearPT 
         BackColor       =   &H00D3C9C0&
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   750
         Width           =   255
      End
      Begin VB.ComboBox cboProductType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   705
         Width           =   2625
      End
      Begin VB.ComboBox cboSection 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1365
         Sorted          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "cboSection"
         Top             =   1080
         Width           =   2640
      End
      Begin VB.TextBox txtRecsFound 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   6000
         Width           =   795
      End
      Begin VB.TextBox txtmaxnum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   5670
         Width           =   795
      End
      Begin VB.CheckBox chkCopies 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Stock on hand"
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
         Height          =   315
         Left            =   2730
         TabIndex        =   1
         Top             =   5250
         Width           =   1350
      End
      Begin VB.CommandButton cmdsearch 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Search"
         Height          =   630
         Left            =   2745
         Picture         =   "frmBrowseSingles.frx":5359
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   5670
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "give or take"
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
         Height          =   210
         Left            =   1710
         TabIndex        =   43
         Top             =   4680
         Width           =   885
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Description/Code"
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
         Left            =   60
         TabIndex        =   40
         Top             =   315
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Max price"
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
         Height          =   210
         Left            =   2940
         TabIndex        =   38
         Top             =   4680
         Width           =   750
      End
      Begin VB.Label lblMeasurement 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1365
         TabIndex        =   36
         Top             =   4935
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
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
         Left            =   45
         TabIndex        =   35
         Top             =   5250
         Width           =   750
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Length"
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
         Left            =   30
         TabIndex        =   34
         Top             =   4950
         Width           =   750
      End
      Begin VB.Label Label3 
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
         Left            =   90
         TabIndex        =   31
         Top             =   750
         Width           =   1185
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
         Left            =   105
         TabIndex        =   30
         Top             =   4275
         Width           =   1185
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
         Left            =   105
         TabIndex        =   29
         Top             =   3930
         Width           =   1185
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
         Left            =   105
         TabIndex        =   28
         Top             =   3585
         Width           =   1185
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
         Left            =   105
         TabIndex        =   27
         Top             =   3225
         Width           =   1185
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
         Left            =   105
         TabIndex        =   26
         Top             =   2880
         Width           =   1185
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
         Left            =   105
         TabIndex        =   25
         Top             =   2535
         Width           =   1185
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
         Left            =   105
         TabIndex        =   24
         Top             =   2175
         Width           =   1185
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
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   1830
         Width           =   1185
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Left            =   540
         TabIndex        =   12
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Found"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   660
         TabIndex        =   8
         Top             =   6045
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00D3D3CB&
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   855
         TabIndex        =   4
         Top             =   5700
         Width           =   390
      End
   End
End
Attribute VB_Name = "frmBrowseSingles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strthing As String
Dim tlkeys As z_TextList
Private oSearchEngine As z_SearchEngineS
Dim colList As Collection
Dim intShowCopies As Integer
Dim lslist As ListItem
Dim roProduct As a_Product
Dim enSource As enProductDataSource
Dim mnu As Menu
Dim XA As New XArrayDB
Dim XN As New XArrayDB
Dim strTime As String
Dim tlCats As z_TextList
Dim BookmarkPointer As Long
Dim tlProductCategorizations As New z_TextList
Dim tlCollection As Collection
Dim tlSuppliers As z_TextList
Dim bWithCopies As Boolean
Dim mWidth As Double
Dim mLength As Double
Dim mPrice As Double
Dim mImage() As Byte
Dim bytTemp() As Byte
Dim mDescriptionOrCode As String
Dim dblLengthMargin As Double
Dim dblWidthMargin As Double
Dim mLengthMargin As Long
Dim mWidthMargin As Long
Dim xMLDoc As ujXML

Private Sub cmdPrint_Click()
    On Error GoTo errHandler
    ExportToXML
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.cmdPrint_Click", , EA_NORERAISE
    HandleError
End Sub
Public Function ExportToXML() As Boolean
    On Error GoTo errHandler
Dim oTF As New z_TextFile
Dim strPath As String
Dim strBillto As String
Dim strDelto As String
Dim strFOFile As String
Dim strFilename As String
Dim strXML As String
Dim strCommand As String
Dim i As Integer
Dim strHTML As String
Dim fs As New FileSystemObject
Dim objXSL As New MSXML2.DOMDocument60
Dim opXMLDOC As New MSXML2.DOMDocument60
Dim objXMLDOC  As New MSXML2.DOMDocument60
Dim strExecutable As String

    Set xMLDoc = New ujXML
    With xMLDoc
        .docProgID = "MSXML2.DOMDocument"
        .docInit "SSI_1"
        .chCreate "SSI"
            .elText = "Selected stock items at " & Format(Now(), "dd/mm/yyyy HH:NN AM/PM")
        
            .elCreateSibling "DetailLine", True
            .chCreate "Col_1"
                .elText = "Code"
            .elCreateSibling "Col_2"
                .elText = "Description"
            .elCreateSibling "Col_3"
                .elText = "Size"
            .elCreateSibling "Col_4"
                .elText = "Price"
            .elCreateSibling "Col_5"
                .elText = "In stock"
                .navUP
            For i = 1 To XN.UpperBound(1)
                If mIsAmongBookmarks(XN, XN.Value(i, 11), GN, 11, "UNIQUEIDENTIFIER") Then
                    .elCreateSibling "DetailLine", True
                    .chCreate "Col_1"
                        .elText = XN.Value(i, 1)
                    .elCreateSibling "Col_2"
                        .elText = XN.Value(i, 2)
                    .elCreateSibling "Col_3"
                        .elText = XN.Value(i, 3)
                    .elCreateSibling "Col_4"
                        .elText = XN.Value(i, 4)
                    .elCreateSibling "Col_5"
                        .elText = XN.Value(i, 5)
                        .navUP
                End If
            Next

        
    End With
    
'FINALLY PRODUCE THE .XML FILE
    strXML = oPC.SharedFolderRoot & "\TEMP\SSI" & ".xml"
    With xMLDoc
        If fs.FileExists(strXML) Then
            fs.DeleteFile strXML
        End If
        .docWriteToFile (strXML), False, "UNICODE", "" 'strHead
    End With

''WRITE THE .RTF FILE
    If Not fs.FileExists(oPC.SharedFolderRoot & "\Templates\SSI_RTF_1.xslt") Then
        MsgBox "You are missing the template file " & "SSI_RTF_1.xslt. Contact Papyrus support." & vbCrLf & "The export is cancelled", vbOKOnly, "Can't do this"
    End If
    objXSL.async = False
    objXSL.ValidateOnParse = False
    objXSL.resolveExternals = False
    strPath = oPC.SharedFolderRoot & "\Templates\SSI_RTF_1.xslt"
    Set fs = New FileSystemObject
    If fs.FileExists(strPath) Then
        objXSL.Load strPath
    End If

    strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1.RTF"
    i = 0
    Do Until fs.FileExists(strFilename) = False
        i = i + 1
        strFilename = oPC.SharedFolderRoot & "\TEMP\SSI_1" & "_" & CStr(i) & ".RTF"
    Loop
    oTF.OpenTextFileToAppend strFilename
    oTF.WriteToTextFile xMLDoc.docObject.transformNode(objXSL)
    oTF.CloseTextFile

    strExecutable = GetPDFExecutable(strFilename)
          If strExecutable = "" Then
              MsgBox "There is no application set on this computer to open the file: " & strFilename & ". The document cannot be displayed", vbOKOnly, "Can't do this"
          Else
              Shell strExecutable & " " & strFilename, vbNormalFocus
          End If
    
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseProducts.ExportToXML"
End Function

Private Sub GN_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal Col As Integer, ByVal CellStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo errHandler
    If Col <> 5 Then Exit Sub
    If CStr(XN(Bookmark, 10)) = "" Then Exit Sub
    bytTemp = XN(Bookmark, 10)
    If UBound(bytTemp) > 0 Then
        CellStyle.Alignment = dbgLeft
        
        CellStyle.ForegroundPicturePosition = dbgFPPictureOnly
        
        CellStyle.ForegroundPicture = ArrayToPictureB(bytTemp(), 0, UBound(bytTemp) + 1)
    End If

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GN_FetchCellStyle(Condition,Split,Bookmark,Col,CellStyle)", _
         Array(Condition, Split, Bookmark, Col, CellStyle), EA_NORERAISE
    HandleError
End Sub

Private Sub GN_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error GoTo errHandler
   If Button = 2 Then   ' Check if right mouse button
                        ' was clicked.
      
      PopupMenu Forms(0).mnuFindForm   ' Display the File menu as a
                        ' pop-up menu.
   End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GN_MouseDown(Button,Shift,x,Y)", Array(Button, Shift, x, Y), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub SetMenu()
    On Error GoTo errHandler

    Forms(0).mnuSaveColumnWidths.Enabled = True
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.SetMenu"
End Sub
Public Sub UnsetMenu()
    On Error GoTo errHandler

    Forms(0).mnuSaveColumnWidths.Enabled = False
      
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.UnsetMenu"
End Sub
Private Sub cboProductType_DblClick()
    On Error GoTo errHandler
    cboProductType = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.cboProductType_DblClick", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboProductType_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
        mSetfocus GN
    End If
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.cboProductType_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub cboSection_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
        mSetfocus GN
    End If
    Exit Sub
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.cboSection_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub

Private Sub chkCopies_Click()
    On Error GoTo errHandler
    oSearchEngine.instock chkCopies
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.chkCopies_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub search(pSearchType As enSearchType, pCriteria As String, Optional pSection As String, Optional pProductType As String)
    On Error GoTo errHandler
Dim strParsedCriteria As String
Dim lngRecsFound As Long
Dim lngResult As Long
Dim lngrows As Long
Dim strArticle As String
Dim strNet As String
Dim strTypes As String
Dim lngSectionID As Long
Dim lngProductTypeID As Long

    strTypes = ""
    
    txtRecsFound = ""
    lngSectionID = 0
    lngProductTypeID = 0
    
    StripArticle pCriteria, strArticle, strNet
    pCriteria = strNet
    oSearchEngine.prisearch
    enSource = enLocalDB
    '--------------
    oSearchEngine.SetupSQLwoCriteria2 False, 200, strTypes    '"NGM"
    
    If pSearchType = enSearchByCatalogue Then
        oSearchEngine.selectcriteria "Catalogue", pCriteria, lngRecsFound
    ElseIf pSearchType = enSearchNormal Then
        oSearchEngine.SimpleSearch pCriteria, lngRecsFound
    Else
        enSource = enLocalDB
        If pSection <> "<ALL>" Then
            lngSectionID = oPC.Configuration.Sections.Key(pSection)
        End If
        If pProductType <> "<ALL>" Then
            lngProductTypeID = oPC.Configuration.ProductTypes.Key(pProductType)
        End If
        oSearchEngine.AdvancedSearch lngRecsFound, pCriteria, lngSectionID, lngProductTypeID
    End If
    'If lngRecsFound > CLng(txtmaxnum) Then MsgBox "Too many records to return, refine your search.", vbInformation + vbOKOnly, "Search result"
    oSearchEngine.execute IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500)
    Set colList = Nothing
    Set colList = oSearchEngine.getcols
    lngrows = oSearchEngine.rows
    txtRecsFound = lngRecsFound
    LoadGrid
    If colList.Count = 0 Then
        Select Case enSource
        Case enLocalDB
            XN.ReDim 1, 1, 1, 12
            XN(1, 1) = "No records"
            GN.ReBind
        End Select
    End If
    If CLng(txtRecsFound) > CLng(txtmaxnum) Then
        MsgBox "No. of records exceeds maximum, please narrow down the search criteria.", , "Criteria too broad"
        Me.GN.Refresh
    End If
    '--------------
    oPC.DisconnectDBShort
    '--------------
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.Search(pSearchType,pCriteria,pSection,pProductType)", Array(pSearchType, _
         pCriteria, pSection, pProductType)
End Sub
Private Sub LoadGrid()
    On Error GoTo errHandler
Dim i As Long

    GN.Splits(0).Columns(5).FetchStyle = True
    Select Case enSource
    Case enLocalDB
        GN.Visible = True
        XN.Clear
        XN.ReDim 1, colList.Count, 1, 14
        For i = 1 To colList.Count
                XN.Value(i, 1) = colList.Item(i).CodeF
                XN.Value(i, 2) = colList.Item(i).Title
                XN.Value(i, 3) = colList.Item(i).LengthandWidth
                XN.Value(i, 4) = colList.Item(i).LocalPriceF
                XN.Value(i, 5) = colList.Item(i).QtyOnHand
                XN.Value(i, 10) = colList.Item(i).Img
                XN.Value(i, 11) = colList.Item(i).PID
                XN.Value(i, 12) = colList.Item(i).EAN
                XN.Value(i, 13) = colList.Item(i).LocalPrice
                XN.Value(i, 14) = colList.Item(i).ImageFilename
        Next
        XN.QuickSort 1, XN.UpperBound(1), 2, XORDER_ASCEND, XTYPE_STRING
        GN.Array = XN
       ' GN.Split(0).Columns(5).FetchCellStyle
        Me.GN.ReBind
        
        
        
    End Select
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.LoadGrid"
End Sub
Public Sub PrintPickingSlip()
    Dim ar As New arPrintBarcodeList2
Dim x As New XArrayDB
Dim i As Integer
Dim j As Integer
Dim f As frmPickingNote
Dim sNote As String
    Set f = New frmPickingNote
    f.Show vbModal
    sNote = f.Note
    Unload f
    x.ReDim 1, GN.SelBookmarks.Count, 1, 14
    For i = 1 To GN.SelBookmarks.Count
        For j = 1 To 14
            x(i, j) = XN(GN.SelBookmarks(i - 1), j)
        Next
    Next

    ar.component x, sNote
    ar.Show vbModal
    Set ar = Nothing

End Sub

Private Sub cmdClearPT_Click()
    On Error GoTo errHandler
    cboProductType = "<ALL>"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.cmdClearPT_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub cmdClearSection_Click()
    On Error GoTo errHandler
    cboSection = "<ALL>"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.cmdClearSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Public Sub mnuSaveLayout()
    On Error Resume Next
    SaveLayout Me.GN, Me.Name, Me.Height, Me.Width
   ' SaveSetting "PBKS", Me.Name, "Formwidth", Me.Width
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.mnuSaveLayout"
End Sub



Private Sub Form_Activate()
    On Error GoTo errHandler
    SetMenu
    XA.Clear
    XA.ReDim 1, 1, 1, 7
    bWithCopies = False
    chkCopies = IIf(bWithCopies, 1, 0)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.Form_Activate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Deactivate()
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.Form_Deactivate", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Initialize()
    On Error GoTo errHandler
Dim i As Integer
    Set oSearchEngine = New z_SearchEngineS
    Set colList = New Collection
    Me.txtLengthMargin = "0.1"
    Me.txtWidthMargin = "0.1"
    dblLengthMargin = 0.1
    dblWidthMargin = 0.1
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormGS", CStr(i), GN.Columns(i - 1).Width)
    Next
    XA.ReDim 1, 1, 1, 7
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.Form_Initialize", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    If Me.WindowState <> 2 Then
        Me.TOP = 20
        Me.Left = 50
    End If
    SetGridLayout Me.GN, Me.Name
    SetFormSize Me
    Set tlSuppliers = New z_TextList
    tlSuppliers.Load ltSupplier, ""
    
'    For i = 1 To GN.Columns.Count
'        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormA", CStr(i), GN.Columns(i - 1).Width)
'    Next
'
    LoadCombo cboSection, oPC.Configuration.Sections
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    Me.cboSection = "<ALL>"
    Me.cboProductType = "<ALL>"
    

 '   GN.Columns(3).Caption = "Supplier"
    txtmaxnum = 50
    
    For i = 1 To GN.Columns.Count
        GN.Columns(i - 1).Width = GetSetting("PBKS", "SearchFormGS", CStr(i), GN.Columns(i - 1).Width)
    Next
    
    Set tlCollection = New Collection
    For i = 1 To 8
        tlCollection.Add New z_TextList
    Next
    For i = 1 To 8
        Me.cboSearch(i - 1).Visible = False
        Me.Label2(i - 1).Visible = False
    Next

    Set tlProductCategorizations = New z_TextList
    tlProductCategorizations.Load ltProductCategorizations
    For i = 0 To tlProductCategorizations.Count - 1
        LoadCombo cboSearch(i), GetTextList(i)
        Me.cboSearch(i).Visible = True
        Me.Label2(i).Visible = True
        Me.Label2(i).Caption = tlProductCategorizations.ItemByOrdinalIndex(i + 1)
        cboSearch(i).text = tlCollection(i + 1).ItemByOrdinalIndex(1)
    Next
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Function GetTextList(i As Integer) As z_TextList
    On Error GoTo errHandler
        tlCollection(i + 1).Load ltProductCategorizationValues, CStr(tlProductCategorizations.KeyByOrdinalIndex(i + 1)), "<ANY>"
        Set GetTextList = tlCollection(i + 1)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GetTextList(i)", i
End Function
Private Sub cmdSearch_Click()
    On Error GoTo errHandler
Dim strSQL As String
Dim i As Integer

    strSQL = ""
    If mDescriptionOrCode > "" Then
        If IsISBN13(mDescriptionOrCode) Then
            search enSearchNormal, mDescriptionOrCode
            Exit Sub
        ElseIf IsHashCode(mDescriptionOrCode) Then
            search enSearchNormal, mDescriptionOrCode
             Exit Sub
       Else
            mDescriptionOrCode = Replace(Me.txtDescriptionOrCode, "'''", "'")
            mDescriptionOrCode = Replace(mDescriptionOrCode, "''", "'")
            If Left(mDescriptionOrCode, 1) = "/" Then
                mDescriptionOrCode = Right(mDescriptionOrCode, Len(mDescriptionOrCode) - 1)
                strSQL = " P_TITLE LIKE '%" & mDescriptionOrCode & "%' AND "
            Else
                strSQL = " P_TITLE LIKE '%" & mDescriptionOrCode & "%' AND "
            End If
        End If
    End If

    For i = 1 To 8
        If cboSearch(i - 1).text <> "<ANY>" And cboSearch(i - 1).text <> "" Then
            strSQL = strSQL & " PATINDEX('%" & tlCollection(i).Key(cboSearch(i - 1).text) & "%',dbo.FlattenCategorization(P_ID)) > 0 AND "
        End If
    Next
    If Right(strSQL, 5) = " AND " Then
        strSQL = Left(strSQL, Len(strSQL) - 5)
    End If

    If mLength > 0 Then
        mLengthMargin = dblLengthMargin * 1000
        mWidthMargin = dblWidthMargin * 1000
        strSQL = strSQL & " AND P_LENGTH <= " & CStr(mLength + mLengthMargin) & " AND P_LENGTH > " & CStr(IIf(mLength - mLengthMargin > 0, mLength - mLengthMargin, 0))
    End If
    If mWidth > 0 Then
        strSQL = strSQL & " AND P_WIDTH <= " & CStr(mWidth + mWidthMargin) & " AND P_WIDTH > " & CStr(IIf(mWidth - mWidthMargin > 0, mWidth - mWidthMargin, 0))
    End If
    
    If mPrice > 0 Then
        strSQL = strSQL & " AND P_SP <= " & CStr(mPrice * oPC.Configuration.DefaultCurrency.Divisor)
    End If
    
    search enSearchAdvanced, strSQL, Me.cboSection, Me.cboProductType


Exit Sub

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.cmdSearch_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub Form_Resize()
        On Error Resume Next
Dim lngDiff As Long
    GN.Width = Me.Width - (GN.Left + 400)
    lngDiff = GN.Height
    GN.Height = Me.Height - (GN.TOP + 1220)
    lngDiff = GN.Height - lngDiff
    cmdClose.TOP = cmdClose.TOP + lngDiff
    cmdPrint.TOP = cmdPrint.TOP + lngDiff
    cmdClose.Left = NonNegative_Lng(Me.Width - 1500)

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.Form_Resize", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Terminate()
    On Error GoTo errHandler
    Set oSearchEngine = Nothing
    Set roProduct = Nothing
    Set colList = Nothing
    Set tlkeys = Nothing
    Set lslist = Nothing
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.Form_Terminate", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    UnsetMenu
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub GN_Click()
    On Error GoTo errHandler
Dim str As String
    If IsNull(GN.Bookmark) Then Exit Sub
    str = FNS(XN.Value(GN.Bookmark, 12))
    If str = "" Then Exit Sub
    On Error Resume Next: Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GN_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub GN_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error GoTo errHandler
Dim str As String
    If IsNull(GN.Bookmark) Then Exit Sub
    str = FNS(XN.Value(GN.Bookmark, 12))
    If str = "" Then Exit Sub
    On Error Resume Next: Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GN_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub


Private Sub GN_DblClick()
    
Dim frmNB As frmProductSinglePreview
Dim lngprod As Long
Dim str As String

On Error Resume Next
    If XN.UpperBound(1) = 0 Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
    If Err Then Exit Sub
    
On Error GoTo errHandler
    BookmarkPointer = GN.Bookmark
    str = FNS(XN.Value(GN.Bookmark, 11))
    If str = "" Then Exit Sub
    Set roProduct = New a_Product
    WaitMsg "Loading . . .", True, Me
    roProduct.Load str, 0, "", strTime
    If roProduct.PID = "" Then Exit Sub
    
    Set frmNB = New frmProductSinglePreview
    frmNB.component roProduct, strTime
    frmNB.Show

    Set roProduct = Nothing
    WaitMsg "", False, Me
    Exit Sub
errHandler:
     ErrPreserve
    If Err.Number = -2147217407 Then   'Access violation
        errRepeat = errRepeat + 1
        LogSaveToFile "Access violation in frmBrowseSingles: GN_DblClick"  'unknown source
        If errRepeat < 5 Then
            Resume Next
        Else
            LogSaveToFile "Access violation in frmBrowseSingles: GN_DblClick after 5 re-attempts"
            MsgBox "Memory error trying to load product form. Please close any other unnecessary applications before trying again.", vbCritical + vbOKOnly, "Can't load product record."
            Err.Clear
            Exit Sub
        End If
    End If
    
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GN_DblClick", , EA_NORERAISE
    HandleError
End Sub
Private Sub GN_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo errHandler
Static Direction As Variant

    If Direction = 0 Then
        Direction = 1
    Else
        Direction = 0
    End If
    If ColIndex = 0 Then ColIndex = 11
    
        XN.QuickSort XN.LowerBound(1), XN.UpperBound(1), ColIndex + 1, Direction, GetRowType(ColIndex + 1) ', 5, XORDER_DESCEND, XTYPE_STRING 'XTYPE_INTEGER
    
    GN.Refresh
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GN_HeadClick(ColIndex)", ColIndex, EA_NORERAISE
    HandleError
End Sub
Private Function GetRowType(ColIndex As Integer) As Variant
    On Error GoTo errHandler
    Select Case ColIndex
        Case 1, 2, 3, 4, 12
            GetRowType = XTYPE_STRING
        Case 5, 6, 7, 8, 9
            GetRowType = XTYPE_INTEGER
    End Select
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GetRowType(ColIndex)", ColIndex
End Function

Private Sub GN_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    
    If KeyAscii = vbKeyReturn Then
        GN_DblClick
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GN_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub



Public Sub AddToTempList()
    On Error GoTo errHandler
Dim str As String
    str = FNS(XN.Value(GN.Bookmark, 11))
    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
        If XA(XA.UpperBound(1), 1) > "" Then
            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
        End If
        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
        XA(XA.UpperBound(1), 4) = 1
        XA(XA.UpperBound(1), 5) = 0
        XA(XA.UpperBound(1), 6) = ""
        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
    End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.AddToTempList"
End Sub
Public Sub PlaceCO()
    On Error GoTo errHandler
Dim frm As New frmPlaceCO
Dim str As String
    str = FNS(XN.Value(GN.Bookmark, 1))
    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
        If XA(XA.UpperBound(1), 1) > "" Then
            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
        End If
        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
        XA(XA.UpperBound(1), 4) = 1
        XA(XA.UpperBound(1), 5) = 0
        XA(XA.UpperBound(1), 6) = ""
        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
    End If
    frm.component XA, "ORDER"
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.PlaceCO"
End Sub
Public Sub PlaceOnReserve()
    On Error GoTo errHandler
Dim frm As New frmPlaceCO
Dim str As String
    str = FNS(XN.Value(GN.Bookmark, 11))
    If XA.Find(1, 4, str) < XA.LowerBound(1) Then
        If XA(XA.UpperBound(1), 1) > "" Then
            XA.ReDim 1, XA.UpperBound(1) + 1, 1, 7
        End If
        XA(XA.UpperBound(1), 1) = FNS(XN.Value(GN.Bookmark, 1))
        XA(XA.UpperBound(1), 2) = FNS(XN.Value(GN.Bookmark, 2))
        XA(XA.UpperBound(1), 3) = FNS(XN.Value(GN.Bookmark, 3))
        XA(XA.UpperBound(1), 4) = 1
        XA(XA.UpperBound(1), 5) = 0
        XA(XA.UpperBound(1), 6) = ""
        XA(XA.UpperBound(1), 7) = FNS(XN.Value(GN.Bookmark, 11))
    End If
    frm.component XA, "RESERVE"
    frm.Show vbModal
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.PlaceOnReserve"
End Sub
Public Sub StartNewList()
    On Error GoTo errHandler
    XA.Clear
    XA.ReDim 1, 1, 1, 9
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.StartNewList"
End Sub


Private Sub txtcritvalues_KeyPress(KeyAscii As Integer)
    On Error GoTo errHandler
    If KeyAscii = vbKeyReturn Then
        cmdSearch_Click
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtcritvalues_KeyPress(KeyAscii)", KeyAscii, EA_NORERAISE
    HandleError
End Sub


Private Sub txtDescriptionOrCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    mDescriptionOrCode = FNS(txtDescriptionOrCode)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtDescriptionOrCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtLength_LostFocus()
    On Error GoTo errHandler
    txtLength = DimensionsF(mLength)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtLength_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtLength_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = (Not IsNumeric(txtLength)) And txtWidth <> ""
    mLength = ConvertDimensionsforStoring(FNDBL(txtLength))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtLength_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtLengthMargin_Change()
    If IsNumeric(txtLengthMargin) Then
        dblLengthMargin = CDbl(txtLengthMargin)
        txtLengthMargin.ForeColor = &H8000000D
    Else
        txtLengthMargin.ForeColor = vbRed
    End If
End Sub
Private Sub txtWidthMargin_Change()
    If IsNumeric(txtWidthMargin) Then
        dblWidthMargin = CDbl(txtWidthMargin)
        txtWidthMargin.ForeColor = &H8000000D
    Else
        txtWidthMargin.ForeColor = vbRed
    End If
End Sub

Private Sub txtWidth_LostFocus()
    On Error GoTo errHandler
    txtWidth = DimensionsF(mWidth)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtWidth_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtWidth_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = (Not IsNumeric(txtWidth)) And txtWidth <> ""
    mWidth = ConvertDimensionsforStoring(FNDBL(txtWidth))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtWidth_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_LostFocus()
    On Error GoTo errHandler
    txtPrice = Format(mPrice, oPC.Configuration.DefaultCurrency.FormatString)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtPrice_LostFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = (Not IsNumeric(txtPrice)) And txtPrice <> ""
    If Not Cancel Then mPrice = FNDBL(txtPrice)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Sub mnuCreateInvoice()
    On Error GoTo errHandler
Dim frm As New frmPlaceTransaction
Dim str As String
Dim TOP As Integer
Dim i As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub

    XA.ReDim 1, GN.SelBookmarks.Count, 1, 9
    For i = 1 To GN.SelBookmarks.Count
        XA(i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        XA(i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(i, 4) = 1
        XA(i, 5) = 0
        XA(i, 6) = "" 'FNS(XN.Value(GN.SelBookmarks(i - 1), 4))
        XA(i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 4))
        XA(i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
        XA(i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 13))
    Next

    frm.component XA, "INVOICE"
    frm.Show 'vbModal
    StartNewList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.mnuCreateInvoice"
End Sub
Public Sub mnuCreateQuotation()
    On Error GoTo errHandler
Dim frm As New frmPlaceTransaction
Dim str As String
Dim TOP As Integer
Dim i As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub

    XA.ReDim 1, GN.SelBookmarks.Count, 1, 9
    For i = 1 To GN.SelBookmarks.Count
        XA(i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        XA(i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(i, 4) = 1
        XA(i, 5) = 0
        XA(i, 6) = "" 'FNS(XN.Value(GN.SelBookmarks(i - 1), 4))
        XA(i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 4))
        XA(i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
        XA(i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 13))
    Next

    frm.component XA, "QUOTATION"
    frm.Show 'vbModal
    StartNewList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.mnuCreateQuotation"
End Sub

Public Property Get NextPID() As String
    On Error GoTo errHandler
    If GN.Array Is Nothing Then
        NextPID = ""
        Exit Property
    End If
    If BookmarkPointer < GN.Array.UpperBound(1) Then
        BookmarkPointer = BookmarkPointer + 1
        NextPID = FNS(XN.Value(BookmarkPointer, 11))
    Else
        NextPID = ""
    End If
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.NextPID"
End Property
Public Property Get PrevPID() As String
    On Error GoTo errHandler
    If GN.Array Is Nothing Then
        PrevPID = ""
        Exit Property
    End If
    If BookmarkPointer > 1 Then
        BookmarkPointer = BookmarkPointer - 1
        PrevPID = FNS(XN.Value(BookmarkPointer, 11))
    Else
        PrevPID = ""
    End If
        
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.PrevPID"
End Property


