VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmProductSingles 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product master (general stock) "
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9465
   ControlBox      =   0   'False
   Icon            =   "frmProductSingles.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleMode       =   0  'User
   ScaleWidth      =   12485.75
   Begin MSComDlg.CommonDialog CD1 
      Left            =   75
      Top             =   6225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImageLoad 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Load"
      Height          =   255
      Left            =   7005
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5055
      Width           =   945
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
      Left            =   1380
      TabIndex        =   51
      Top             =   2055
      Width           =   2550
   End
   Begin VB.TextBox txtLength 
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
      Left            =   1365
      TabIndex        =   48
      Top             =   4680
      Width           =   1380
   End
   Begin VB.TextBox txtWidth 
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
      Left            =   1365
      TabIndex        =   47
      Top             =   5055
      Width           =   1380
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   0
      ItemData        =   "frmProductSingles.frx":030A
      Left            =   6360
      List            =   "frmProductSingles.frx":030C
      TabIndex        =   39
      Top             =   1830
      Width           =   2940
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      ItemData        =   "frmProductSingles.frx":030E
      Left            =   6360
      List            =   "frmProductSingles.frx":0310
      TabIndex        =   38
      Top             =   2190
      Width           =   2940
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   2
      ItemData        =   "frmProductSingles.frx":0312
      Left            =   6360
      List            =   "frmProductSingles.frx":0314
      TabIndex        =   37
      Top             =   2550
      Width           =   2940
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   3
      ItemData        =   "frmProductSingles.frx":0316
      Left            =   6360
      List            =   "frmProductSingles.frx":0318
      TabIndex        =   36
      Top             =   2925
      Width           =   2940
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   4
      ItemData        =   "frmProductSingles.frx":031A
      Left            =   6360
      List            =   "frmProductSingles.frx":031C
      TabIndex        =   35
      Top             =   3285
      Width           =   2940
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   5
      ItemData        =   "frmProductSingles.frx":031E
      Left            =   6360
      List            =   "frmProductSingles.frx":0320
      TabIndex        =   34
      Top             =   3645
      Width           =   2940
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   6
      ItemData        =   "frmProductSingles.frx":0322
      Left            =   6360
      List            =   "frmProductSingles.frx":0324
      TabIndex        =   33
      Top             =   4005
      Width           =   2940
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   7
      ItemData        =   "frmProductSingles.frx":0326
      Left            =   6360
      List            =   "frmProductSingles.frx":0328
      TabIndex        =   32
      Top             =   4365
      Width           =   2940
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
      Left            =   8295
      Picture         =   "frmProductSingles.frx":032A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7065
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
      Left            =   7275
      Picture         =   "frmProductSingles.frx":06B4
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7065
      Width           =   1000
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
      Left            =   4740
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8400
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Identification codes"
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
      Height          =   1365
      Left            =   240
      TabIndex        =   23
      Top             =   120
      Width           =   9000
      Begin VB.TextBox Text1 
         BackColor       =   &H00D3D3CB&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
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
         Height          =   840
         Left            =   3135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "frmProductSingles.frx":0A3E
         Top             =   390
         Width           =   5415
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
         Left            =   1110
         TabIndex        =   1
         Top             =   810
         Width           =   1680
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
         Left            =   1110
         TabIndex        =   0
         Top             =   420
         Width           =   1680
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
         Left            =   390
         TabIndex        =   25
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label11 
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
         Left            =   180
         TabIndex        =   24
         Top             =   465
         Width           =   870
      End
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
      Left            =   1365
      TabIndex        =   12
      Top             =   4185
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
      Left            =   1365
      TabIndex        =   11
      Top             =   3795
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Status"
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
      Height          =   1305
      Left            =   465
      TabIndex        =   7
      Top             =   5565
      Width           =   2280
      Begin VB.OptionButton optRP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Sold"
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
         Height          =   270
         Left            =   270
         TabIndex        =   10
         Top             =   945
         Width           =   1575
      End
      Begin VB.OptionButton optOOP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Reserved"
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
         Height          =   270
         Left            =   270
         TabIndex        =   9
         Top             =   630
         Width           =   1575
      End
      Begin VB.OptionButton optIP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Available"
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
         Height          =   270
         Left            =   270
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   1575
      End
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
      Left            =   2280
      TabIndex        =   13
      Top             =   7050
      Visible         =   0   'False
      Width           =   1380
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
      Left            =   1365
      TabIndex        =   6
      Top             =   1665
      Width           =   2565
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   4155
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   6855
      Visible         =   0   'False
      Width           =   2625
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
      Left            =   1995
      TabIndex        =   5
      Top             =   7905
      Visible         =   0   'False
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
      Left            =   1995
      TabIndex        =   4
      Top             =   7470
      Visible         =   0   'False
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
      Left            =   1365
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3120
      Width           =   3225
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
      Left            =   1365
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2490
      Width           =   3225
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2625
      Left            =   3390
      Top             =   5055
      Width           =   3585
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      Left            =   510
      TabIndex        =   52
      Top             =   2100
      Width           =   750
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Length"
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
      Left            =   555
      TabIndex        =   50
      Top             =   4695
      Width           =   750
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
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
      Left            =   555
      TabIndex        =   49
      Top             =   5055
      Width           =   750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock group"
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
      Index           =   1
      Left            =   4710
      TabIndex        =   46
      Top             =   2198
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock group"
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
      Index           =   2
      Left            =   4710
      TabIndex        =   45
      Top             =   2566
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock group"
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
      Index           =   3
      Left            =   4710
      TabIndex        =   44
      Top             =   2934
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock group"
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
      Index           =   4
      Left            =   4710
      TabIndex        =   43
      Top             =   3302
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock group"
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
      Index           =   5
      Left            =   4710
      TabIndex        =   42
      Top             =   3670
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock group"
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
      Index           =   6
      Left            =   4710
      TabIndex        =   41
      Top             =   4038
      Width           =   1545
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Stock group"
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
      Index           =   7
      Left            =   4710
      TabIndex        =   40
      Top             =   4410
      Width           =   1545
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
      Left            =   1830
      TabIndex        =   29
      Top             =   8370
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Stock group"
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
      Index           =   0
      Left            =   4710
      TabIndex        =   28
      Top             =   1890
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   540
      TabIndex        =   22
      Top             =   4185
      Width           =   750
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   555
      TabIndex        =   21
      Top             =   3810
      Width           =   750
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   1110
      TabIndex        =   20
      Top             =   7095
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   225
      TabIndex        =   19
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   1275
      TabIndex        =   17
      Top             =   7935
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   660
      TabIndex        =   16
      Top             =   7515
      Visible         =   0   'False
      Width           =   1260
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
      Left            =   645
      TabIndex        =   15
      Top             =   3135
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
      Left            =   75
      TabIndex        =   14
      Top             =   2535
      Width           =   1215
   End
End
Attribute VB_Name = "frmProductSingles"
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
Dim tlProductCategorizations As New z_TextList
Dim tlCollection As Collection
Dim tlSuppliers As z_TextList

Sub component(pProduct As a_Product, Optional pPrevForm As Form)
    On Error GoTo errHandler
    Set frmPrevious = pPrevForm
    Set oProd = pProduct
    oProd.BeginEdit
    If oProd.IsNew Then
        oProd.SetGeneralProduct
    End If
    oProd.GetStatus
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.component(pProduct,pPrevForm)", Array(pProduct, pPrevForm)
End Sub

Private Sub cboProductType_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetProductTypeID oPC.Configuration.ProductTypes.key(cboProductType)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cboProductType_Click", , EA_NORERAISE
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
    ErrorIn "frmProductSingles.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cboSearch_Change(Index As Integer)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
   ' oProd.ProductCategories(Index).    'oPC.Configuration.ProductTypes.key (cboProductType)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cboSearch_Change(Index)", Index, EA_NORERAISE
    HandleError
End Sub

Private Sub cboSearch_Click(Index As Integer)
    On Error GoTo errHandler
Dim oProdCat As a_Product_Category
    If flgLoading Then Exit Sub
    
    Set oProdCat = oProd.ProductCategories.ItemByCatID(tlProductCategorizations.KeyByOrdinalIndex(Index + 1))
    If oProdCat Is Nothing And cboSearch(Index) <> "" And cboSearch(Index) <> "<n/a>" Then
        Set oProdCat = oProd.ProductCategories.Add
        oProdCat.CatID = tlProductCategorizations.KeyByOrdinalIndex(Index + 1)
    Else
        If Not oProdCat Is Nothing Then
            If cboSearch(Index) = "" Or cboSearch(Index) = "<n/a>" Then
                oProdCat.Delete
                oProdCat.ApplyEdit
                oProdCat.BeginEdit
                Exit Sub
            Else
                oProdCat.CatValueID = tlCollection(Index + 1).key(cboSearch(Index))
                oProdCat.ApplyEdit
                oProdCat.BeginEdit
            End If
        Else
            Exit Sub
        End If
    End If
    
    oProdCat.CatValueID = tlCollection(Index + 1).key(cboSearch(Index))
    oProdCat.Description = cboSearch(Index)
    oProdCat.ApplyEdit
    oProdCat.BeginEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cboSearch_Click(Index)", Index, EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdAddSection_Click()
'    On Error GoTo errHandler
'Dim oPSEC As New a_ProductSection
'    If flgLoading Then Exit Sub
'    If cboSection = "" Then Exit Sub
'    Set oPSEC = oProd.ProductSections.Add
' '   oCC.BeginEdit
'    oPSEC.pID = oProd.pID
'    oPSEC.SECID = oPC.Configuration.Sections.key(cboSection)
'    oPSEC.Description = cboSection
'    If oProd.ProductSections.Count = 0 Then
'        oPSEC.Priority = 99
'    End If
'    oPSEC.ApplyEdit
'    oPSEC.BeginEdit
'    cboSection.RemoveItem cboSection.ListIndex
'    If cboSection.ListCount > 0 Then
'        cboSection.ListIndex = 0
'    Else
'        cboSection.ListIndex = -1
'    End If
'    LoadPSECs
'    Exit Sub
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdAddSection_Click", , EA_NORERAISE
'    HandleError
'End Sub

'Private Sub cmdGenerateEAN_Click()
'Dim oProdCode As New z_ProdCode
'    oProdCode.SetCodesForBook txtCode
'    oProd.SetEAN oProdCode.Ean
'    txtEAN = oProd.Ean
'End Sub


Private Sub cmdChangeType_Click()
    On Error GoTo errHandler
    If MsgBox("You want to change this product to be a book?", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    Else
        oProd.SetProductType "B"
        oProd.ApplyEdit
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cmdChangeType_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMultSP_Click()
    On Error GoTo errHandler
    If IsNumeric(Me.txtSP) Then
        oProd.SetSP CStr(CDbl(txtSP * oPC.Configuration.DefaultCurrency.Divisor) * 1.14)
        txtSP = oProd.SPF
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cmdMultSP_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub cboSection_Click()
    On Error GoTo errHandler
Dim oPS As a_ProductSection
    If flgLoading Then Exit Sub
        oProd.ProductSections.Delete
        Set oPS = oProd.ProductSections.Add
        oPS.PID = oProd.PID
        oPS.SECID = oPC.Configuration.Sections.key(cboSection)
        oPS.Description = cboSection
        oPS.ApplyEdit
        oPS.BeginEdit
        
    'oProd.preoductsection
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cboSection_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdImageLoad_Click()
    On Error GoTo errHandler
Dim strFilePath As String
Dim oSQL As New z_SQL
Dim bytTemp() As Byte

    CD1.ShowOpen
    strFilePath = CD1.FileName
    AddImageToDB strFilePath, oProd.PID
    LoadPictureFromDB oProd.PID
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cmdImageLoad_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadPictureFromDB(PID As String)
    On Error GoTo errHandler
Dim bytTemp() As Byte
    bytTemp = ImageFromDB(PID)
    If UBound(bytTemp) > 0 Then
        Image1.Stretch = True
'        Image1.Width = 3000
'        Image1.Height = 1500
        Set Image1 = ArrayToPictureB(bytTemp(), 0, UBound(bytTemp) + 1)
    Else
        Set Image1.Picture = LoadPicture
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.LoadPictureFromDB(PID)", PID
End Sub

Private Sub txtLength_GotFocus()
    On Error GoTo errHandler
    txtLength = DimensionsF(oProd.Length)
    AutoSelect txtLength
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtLength_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtLength_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oProd.SetLength(txtLength) Then
        Cancel = True
    End If
    txtLength = oProd.LengthF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtLength_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtWidth_GotFocus()
    On Error GoTo errHandler
    txtWidth = DimensionsF(oProd.Width)
    AutoSelect txtWidth
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtWidth_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtWidth_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    If Not oProd.SetWidth(txtWidth) Then
        Cancel = True
    End If
    txtWidth = oProd.WidthF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtWidth_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtSP_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
Dim CtrlDown
    CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyAdd Then
        If CtrlDown Then
            If IsNumeric(Me.txtSP) Then
                oProd.SetSP CStr(CDbl(txtSP) * 1.14)
                txtSP = oProd.SP
            End If
        End If
    End If
'  Dim ShiftDown, AltDown, CtrlDown, Txt
'   ShiftDown = (Shift And vbShiftMask) > 0
'   AltDown = (Shift And vbAltMask) > 0'   CtrlDown = (Shift And vbCtrlMask) > 0
'   If KeyCode = vbKeyAdd Then   ' Display key combinations.
'   If ShiftDown And CtrlDown And AltDown Then
'      Txt = "SHIFT+CTRL+ALT+F2."
'   ElseIf ShiftDown And AltDown Then
'      Txt = "SHIFT+ALT+F2."
'   ElseIf ShiftDown And CtrlDown Then
'      Txt = "SHIFT+CTRL+F2."
'   ElseIf CtrlDown And AltDown Then
'      Txt = "CTRL+ALT+F2."
'   ElseIf ShiftDown Then
'      Txt = "SHIFT+F2."
'   ElseIf CtrlDown Then
'   Txt = "CTRL+F2."
'   ElseIf AltDown Then
'      Txt = "ALT+F2."
'   ElseIf Shift = 0 Then
'      Txt = "F2."
'   End If
'   Text1.Text = "You pressed " & Txt
'   End If
    
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtSP_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSetDefault_Click()
    On Error GoTo errHandler
    Me.txtVAT = oPC.Configuration.VATRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cmdSetDefault_Click", , EA_NORERAISE
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cmdSupplier_Click", , EA_NORERAISE
    HandleError
End Sub

'Private Sub cmdUp_Click()
'   oProd.ProductSections.mark oProd.ProductSections.key(lvw.SelectedItem)
'    LoadPSECs
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oProd.IsEditing Then oProd.CancelEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub oProd_Valid(strMsg As String)
    On Error GoTo errHandler
    Me.txtErrors = strMsg
    Me.cmdok.Enabled = (strMsg = "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.oProd_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oProd.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdNewCode_Click()
    On Error GoTo errHandler
    Me.txtCode = "#"
    oProd.SetCode "#"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cmdNewCode_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim strMsg As String
Dim frmPreview As frmProductSinglePreview
Dim i As Integer
WaitMsg "Saving product . . .", True, Me
    
'    For i = 1 To 8
'        If cboSearch(i).Text <> "" And cboSearch(i).Text <> "<n/a>" Then
'
'
'        End If
'        oProd.CatalogueEntries
'        oPSEC.pID = oProd.pID
'        oPSEC.SECID = oPC.Configuration.Sections.key(cboSection)
'        oPSEC.Description = cboSection
'        If oProd.ProductSections.Count = 0 Then
'            oPSEC.Priority = 99
'        End If
'        oPSEC.ApplyEdit
'        oPSEC.BeginEdit
'
'
'    Next i
    
    
    
    
    oProd.SetBFDistributorCode "XXX"
    oProd.ApplyEdit lngResult, strMsg
    If lngResult = 99 Then
        WaitMsg "", False, Me
        If strMsg = "DUPLICATE" Then
            MsgBox "Invalid values - check that the code is has not been already used", vbInformation, "Save failed"
        ElseIf strMsg = "TIMEOUT" Then
            MsgBox "The operation has timed out. The record is probably locked by another user." & vbCrLf & "Cancel your update or try again. ", vbInformation, "Save failed"
        End If
    Else
        If frmPrevious Is Nothing Then
            Set frmPreview = New frmProductSinglePreview
        Else
            Set frmPreview = frmPrevious
        End If
        frmPreview.component oProd
        frmPreview.RefreshForm
        frmPreview.Show
        WaitMsg "", False, Me
        Unload Me
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
Dim i As Integer
    flgLoading = True
    If Me.WindowState <> 2 Then
        Left = 10
        top = 10
        Width = 10000
        Height = 8800
    End If
    LoadControls
    
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
        If tlCollection(i + 1).ItemByOrdinalIndex(1) > "" Then cboSearch(i).Text = tlCollection(i + 1).ItemByOrdinalIndex(1)
        If Not oProd.ProductCategories.ItemByCatID(CStr(tlProductCategorizations.KeyByOrdinalIndex(i + 1))) Is Nothing Then
            cboSearch(i).Text = oProd.ProductCategories.ItemByCatID(CStr(tlProductCategorizations.KeyByOrdinalIndex(i + 1))).Description
        End If
    Next
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Function GetTextList(i As Integer) As z_TextList
    On Error GoTo errHandler
        tlCollection(i + 1).Load ltProductCategorizationValues, CStr(tlProductCategorizations.KeyByOrdinalIndex(i + 1)), "<n/a>"
        Set GetTextList = tlCollection(i + 1)
    Exit Function
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.GetTextList(i)", i
End Function

Private Sub LoadControls()
    On Error GoTo errHandler
    
    txtCode = oProd.code
    Me.txtEAN = oProd.EAN
    txtTitle = oProd.Title
    txtSubtitle = oProd.SubTitle
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    txtSP = oProd.SPF
    txtCost = oProd.CostF
    txtVAT = oProd.VATRateF
    txtLength = oProd.LengthF
    txtWidth = oProd.WidthF
    LoadCombo cboSection, oPC.Configuration.Sections
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    cboProductType = oPC.Configuration.ProductTypes.Item(oProd.ProductTypeID)
    Select Case oProd.Status
    Case "O"
        optOOP.Value = True
    Case "R"
        optRP.Value = True
    Case Else
        optIP.Value = True
    End Select
    Me.cboSection = oProd.ProductSections(1).Description
    LoadPictureFromDB oProd.PID

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.LoadControls"
End Sub

Private Sub optIP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enInPrint
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.optIP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOOP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enOutOfPrint
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.optOOP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optRP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enAwaitingReprint
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.optRP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Change()
    On Error GoTo errHandler
Dim intPos As Integer
    On Error Resume Next
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
    ErrorIn "frmProductSingles.txtCode_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCode_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    'Cancel = Not oProd.SetCode(txtCode)
    oProd.SetCode txtCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtEAN_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtEAN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtEAN_GotFocus", , EA_NORERAISE
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
    ErrorIn "frmProductSingles.txtEAN_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEAN_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    oProd.SetEAN txtEAN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtEAN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtSP_GotFocus()
    On Error GoTo errHandler
    txtSP = oProd.SP
    AutoSelect txtSP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtSP_GotFocus", , EA_NORERAISE
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
    ErrorIn "frmProductSingles.txtSP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCost_GotFocus()
    On Error GoTo errHandler
    txtCost = oProd.Cost
    AutoSelect txtCost
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtCost_GotFocus", , EA_NORERAISE
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
    ErrorIn "frmProductSingles.txtCost_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtSubtitle_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtSubtitle = oProd.SubTitle
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtSubtitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSubtitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtSubtitle_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmProductSingles.txtSubtitle_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtTitle = oProd.Title
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmProductSingles.txtTitle_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPublisher = oProd.Publisher
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtPublisher_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtPublisher_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmProductSingles.txtPublisher_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtEdition = oProd.Edition
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtEdition_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductSingles.txtEdition_Validate(Cancel)", Cancel, EA_NORERAISE
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
    ErrorIn "frmProductSingles.txtEdition_Change", , EA_NORERAISE
    HandleError
End Sub



