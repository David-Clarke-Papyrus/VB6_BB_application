VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProductNB 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Product master (general stock) "
   ClientHeight    =   8235
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9795
   ControlBox      =   0   'False
   Icon            =   "frmProductNB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8235
   ScaleMode       =   0  'User
   ScaleWidth      =   12921.07
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C4BCA4&
      Caption         =   "+ VAT"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7410
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3120
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C4BCA4&
      Caption         =   "+ VAT"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2385
      Width           =   525
   End
   Begin VB.CommandButton cmdMultSP 
      BackColor       =   &H00C4BCA4&
      Caption         =   "+ VAT"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7410
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   2760
      Width           =   525
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
      Left            =   8040
      Picture         =   "frmProductNB.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   6930
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
      Left            =   7020
      Picture         =   "frmProductNB.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   6930
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
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   4080
      Width           =   570
   End
   Begin VB.CommandButton cmdChangeType 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Change this product type to a book"
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
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   6960
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Sections"
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
      Height          =   1890
      Left            =   510
      TabIndex        =   34
      Top             =   4950
      Width           =   4215
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
         TabIndex        =   39
         Top             =   1410
         Width           =   330
      End
      Begin VB.CommandButton cmdAddSection 
         BackColor       =   &H00C4BCA4&
         Caption         =   "&Add"
         Height          =   315
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   330
         Width           =   750
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
         Left            =   510
         TabIndex        =   18
         Top             =   330
         Width           =   2490
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   930
         Left            =   495
         TabIndex        =   40
         Top             =   780
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1640
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Codes and numbers"
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
      Height          =   1800
      Left            =   270
      TabIndex        =   30
      Top             =   240
      Width           =   8730
      Begin VB.CheckBox chkNonCounted 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Non counted type stock"
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
         Left            =   1110
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
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
         Height          =   990
         Left            =   3135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "frmProductNB.frx":0A1E
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
         TabIndex        =   32
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
         TabIndex        =   31
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
      Left            =   6015
      TabIndex        =   15
      Top             =   3000
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
      Left            =   6015
      TabIndex        =   14
      Top             =   2625
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
      Left            =   6015
      TabIndex        =   13
      Top             =   2250
      Width           =   1380
   End
   Begin VB.CheckBox chkObsolete 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3D3CB&
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
      Left            =   2175
      TabIndex        =   12
      Top             =   7290
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Suppliers' status"
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
      Left            =   6030
      TabIndex        =   8
      Top             =   4920
      Width           =   2280
      Begin VB.OptionButton optRP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "On backorder"
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
         TabIndex        =   11
         Top             =   945
         Width           =   1575
      End
      Begin VB.OptionButton optOOP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Unavailable"
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
         Top             =   630
         Width           =   1575
      End
      Begin VB.OptionButton optIP 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Available"
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
         Top             =   315
         Value           =   -1  'True
         Width           =   1575
      End
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
      Left            =   7470
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3555
      Width           =   1755
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
      Left            =   6000
      TabIndex        =   16
      Top             =   3585
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
      Left            =   1470
      TabIndex        =   7
      Top             =   4500
      Width           =   2565
   End
   Begin VB.TextBox txtErrors 
      Appearance      =   0  'Flat
      BackColor       =   &H00D3D3CB&
      BorderStyle     =   0  'None
      ForeColor       =   &H000000FF&
      Height          =   885
      Left            =   4065
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   6930
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
      Left            =   1485
      TabIndex        =   6
      Top             =   4050
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
      Left            =   1485
      TabIndex        =   5
      Top             =   3615
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
      Left            =   1485
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2880
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
      Left            =   1485
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2250
      Width           =   3225
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
      Left            =   6015
      TabIndex        =   38
      Top             =   4050
      Width           =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   5130
      TabIndex        =   37
      Top             =   4110
      Width           =   810
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
      Left            =   5205
      TabIndex        =   29
      Top             =   3000
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
      Left            =   5205
      TabIndex        =   28
      Top             =   2640
      Width           =   750
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   5205
      TabIndex        =   27
      Top             =   2265
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
      Left            =   4830
      TabIndex        =   26
      Top             =   3630
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
      Left            =   330
      TabIndex        =   25
      Top             =   4545
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
      Left            =   765
      TabIndex        =   23
      Top             =   4080
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
      Left            =   150
      TabIndex        =   22
      Top             =   3660
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
      Left            =   765
      TabIndex        =   21
      Top             =   2895
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
      Left            =   195
      TabIndex        =   20
      Top             =   2295
      Width           =   1215
   End
End
Attribute VB_Name = "frmProductNB"
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
    ErrorIn "frmProductNB.component(pProduct,pPrevForm)", Array(pProduct, pPrevForm)
End Sub


'Private Sub cboSection_Click()
'    If flgLoading Then Exit Sub
'    oProd.SetSection cboSection
'    txtSection = oProd.Section
'End Sub
Private Sub cboProductType_Click()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    oProd.SetProductTypeID oPC.Configuration.ProductTypes.Key(cboProductType)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.cboProductType_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub chkNonCounted_Click()
    On Error GoTo errHandler
    If chkNonCounted Then
        oProd.SetMagsEtc
    Else
        oProd.SetGeneralProduct
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.chkNonCounted_Click", , EA_NORERAISE
    HandleError
End Sub



'Private Sub chkServiceItem_Click()
'    oProd.ServiceItem = chkServiceItem
'End Sub

Private Sub chkObsolete_Click()
    On Error GoTo errHandler
    oProd.Obsolete = chkObsolete
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.chkObsolete_Click", , EA_NORERAISE
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
    ErrorIn "frmProductNB.cmdDelete_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdAddSection_Click()
    On Error GoTo errHandler
Dim oPSEC As New a_ProductSection
    If flgLoading Then Exit Sub
    If cboSection = "" Then Exit Sub
    Set oPSEC = oProd.ProductSections.Add
 '   oCC.BeginEdit
    oPSEC.PID = oProd.PID
    oPSEC.SECID = oPC.Configuration.Sections.Key(cboSection)
    oPSEC.Description = cboSection
    If oProd.ProductSections.Count = 0 Then
        oPSEC.Priority = 99
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
'errHandler:
'    If ErrMustStop Then Debug.Assert False: Resume
'    ErrorIn "frmProduct.cmdAddSection_Click", , EA_NORERAISE
'    HandleError
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.cmdAddSection_Click", , EA_NORERAISE
    HandleError
End Sub

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
    ErrorIn "frmProductNB.cmdChangeType_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdMultSP_Click()
    On Error GoTo errHandler
    If IsNumeric(Me.txtSP) Then
        oProd.SetSP CStr(CDbl(txtSP * oPC.Configuration.DefaultCurrency.Divisor) * (100 + oPC.Configuration.VATRate) / 100)
        txtSP = oProd.SPF
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.cmdMultSP_Click", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSP_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo errHandler
Dim CtrlDown
    CtrlDown = (Shift And vbCtrlMask) > 0
    If KeyCode = vbKeyAdd Then
        If CtrlDown Then
            If IsNumeric(Me.txtSP) Then
                oProd.SetSP CStr(CDbl(txtSP) * (100 + oPC.Configuration.VATRate) / 100)
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
    ErrorIn "frmProductNB.txtSP_KeyDown(KeyCode,Shift)", Array(KeyCode, Shift), EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSetDefault_Click()
    On Error GoTo errHandler
    Me.txtVAT = oPC.Configuration.VATRate
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.cmdSetDefault_Click", , EA_NORERAISE
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
    ErrorIn "frmProductNB.cmdSupplier_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdUP_Click()
    On Error GoTo errHandler
   oProd.ProductSections.Mark oProd.ProductSections.Key(lvw.SelectedItem)
    LoadPSECs
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.cmdUP_Click", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errHandler
    If oProd.IsEditing Then oProd.CancelEdit
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.Form_Unload(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub oProd_Valid(strMsg As String)
    On Error GoTo errHandler
    Me.txtErrors = strMsg
    Me.cmdOK.Enabled = (strMsg = "")
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.oProd_Valid(strMsg)", strMsg, EA_NORERAISE
    HandleError
End Sub
Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    oProd.CancelEdit
    Unload Me
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdNewCode_Click()
    On Error GoTo errHandler
    Me.txtCode = "#"
    oProd.SetCode "#"
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.cmdNewCode_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
Dim lngResult As Long
Dim strMsg As String
Dim frmPreview As frmProductNBPrev

    WaitMsg "Saving product . . .", True, Me
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
            Set frmPreview = New frmProductNBPrev
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
    ErrorIn "frmProductNB.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    If Me.WindowState <> 2 Then
        Left = 10
        TOP = 10
        Width = 10000
        Height = 8800
    End If
    LoadControls
   ' Me.txtEAN.Enabled = left(txtEAN, 1) <> "2" 'only enable standard
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.Form_Load", , EA_NORERAISE
    HandleError
End Sub
Private Sub LoadControls()
    On Error GoTo errHandler
    flgLoading = True
    txtCode = oProd.code
    Me.txtEAN = oProd.EAN
    txtTitle = oProd.Title
    txtSubtitle = oProd.SubTitle
    txtEdition = oProd.Edition
    txtPublisher = oProd.Publisher
    txtRRP = oProd.RRPF
    txtSP = oProd.SPF
    txtCost = oProd.CostF
  '  txtSection = oProd.Section
    Me.txtVAT = oProd.VATRateF
    LoadCombo cboSection, oPC.Configuration.Sections
    LoadCombo cboProductType, oPC.Configuration.ProductTypes
    cboProductType = oPC.Configuration.ProductTypes.Item(oProd.ProductTypeID)
    Me.chkNonCounted = IIf(oProd.IsMagsEtc, 1, 0)
    Me.chkObsolete = IIf(oProd.Obsolete, 1, 0)
    Select Case oProd.Status
    Case "O"
        optOOP.Value = True
    Case "R"
        optRP.Value = True
    Case Else
        optIP.Value = True
    End Select
    LoadPSECs
    flgLoading = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.LoadControls"
End Sub

Private Sub optIP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enInPrint
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.optIP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optOOP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enOutOfPrint
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.optOOP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub optRP_Click()
    On Error GoTo errHandler
    oProd.SetStatus enAwaitingReprint
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.optRP_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtCode_Change()
            On Error Resume Next
Dim intPos As Integer
    mCancel = Not oProd.SetCode(txtCode)
    On Error Resume Next
    If Err Then
      Beep
      intPos = txtCode.SelStart
      txtCode = oProd.code
      txtCode.SelStart = intPos - 1
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtCode_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCode_Validate(Cancel As Boolean)
            On Error Resume Next
    oProd.SetCode txtCode
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtCode_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtEAN_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtEAN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtEAN_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEAN_Change()
            On Error Resume Next
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
    ErrorIn "frmProductNB.txtEAN_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtEAN_Validate(Cancel As Boolean)
            On Error Resume Next
    'Cancel = Not oProd.SetEAN(txtEAN)
    oProd.SetEAN txtEAN
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtEAN_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub


Private Sub txtRRP_GotFocus()
    On Error GoTo errHandler
    txtRRP = oProd.RRP
    AutoSelect txtRRP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtRRP_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtRRP_Validate(Cancel As Boolean)
            On Error Resume Next
    If flgLoading Then Exit Sub
    If Not oProd.SetRRP(txtRRP) Then
        Cancel = True
    End If
    txtRRP = oProd.RRPF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtRRP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

'Private Sub txtSection_Validate(Cancel As Boolean)
'    oProd.SetSectionAll txtSection
'    txtSection = oProd.Section
'End Sub

Private Sub txtSP_GotFocus()
    On Error GoTo errHandler
    txtSP = oProd.SP
    AutoSelect txtSP
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtSP_GotFocus", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtSP_Validate(Cancel As Boolean)
            On Error Resume Next
    If flgLoading Then Exit Sub
    If Not oProd.SetSP(txtSP) Then
        Cancel = True
    End If
    txtSP = oProd.SPF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtSP_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtCost_GotFocus()
    On Error GoTo errHandler
    txtCost = oProd.Cost
    AutoSelect txtCost
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtCost_GotFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtCost_Validate(Cancel As Boolean)
            On Error Resume Next
    If flgLoading Then Exit Sub
    If Not oProd.SetCost(txtCost) Then
        Cancel = True
    End If
    txtCost = oProd.CostF
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtCost_Validate(Cancel)", Cancel, EA_NORERAISE
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
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtSubtitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtSubtitle_Validate(Cancel As Boolean)
            On Error Resume Next
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtSubtitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtSubtitle_Change()
            On Error Resume Next
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
    ErrorIn "frmProductNB.txtSubtitle_Change", , EA_NORERAISE
    HandleError
End Sub


Private Sub txtTitle_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtTitle = oProd.Title
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtTitle_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Validate(Cancel As Boolean)
            On Error Resume Next
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtTitle_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtTitle_Change()
            On Error Resume Next
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
    ErrorIn "frmProductNB.txtTitle_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtPublisher = oProd.Publisher
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtPublisher_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_Validate(Cancel As Boolean)
            On Error Resume Next
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtPublisher_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtPublisher_Change()
            On Error Resume Next
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
    ErrorIn "frmProductNB.txtPublisher_Change", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_LostFocus()
    On Error GoTo errHandler
    If flgLoading Then Exit Sub
    txtEdition = oProd.Edition
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtEdition_LostFocus", , EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_Validate(Cancel As Boolean)
            On Error Resume Next
    Cancel = mCancel
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmProductNB.txtEdition_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
Private Sub txtEdition_Change()
            On Error Resume Next
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
    ErrorIn "frmProductNB.txtEdition_Change", , EA_NORERAISE
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
    ErrorIn "frmProductNB.LoadPSECs"
End Sub

