VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
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
   Begin TrueOleDBGrid60.TDBGrid GN 
      Height          =   5670
      Left            =   4590
      OleObjectBlob   =   "frmBrowseSingles.frx":058A
      TabIndex        =   9
      Top             =   240
      Width           =   11130
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Height          =   615
      Left            =   4605
      Picture         =   "frmBrowseSingles.frx":4B85
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5970
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
         Left            =   1350
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
         Left            =   1350
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
         ItemData        =   "frmBrowseSingles.frx":4F0F
         Left            =   1365
         List            =   "frmBrowseSingles.frx":4F11
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
         ItemData        =   "frmBrowseSingles.frx":4F13
         Left            =   1365
         List            =   "frmBrowseSingles.frx":4F15
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
         ItemData        =   "frmBrowseSingles.frx":4F17
         Left            =   1365
         List            =   "frmBrowseSingles.frx":4F19
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
         ItemData        =   "frmBrowseSingles.frx":4F1B
         Left            =   1365
         List            =   "frmBrowseSingles.frx":4F1D
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
         ItemData        =   "frmBrowseSingles.frx":4F1F
         Left            =   1365
         List            =   "frmBrowseSingles.frx":4F21
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
         ItemData        =   "frmBrowseSingles.frx":4F23
         Left            =   1365
         List            =   "frmBrowseSingles.frx":4F25
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
         ItemData        =   "frmBrowseSingles.frx":4F27
         Left            =   1365
         List            =   "frmBrowseSingles.frx":4F29
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
         ItemData        =   "frmBrowseSingles.frx":4F2B
         Left            =   1365
         List            =   "frmBrowseSingles.frx":4F2D
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
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
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
         Width           =   780
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
         Picture         =   "frmBrowseSingles.frx":4F2F
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   5670
         Width           =   1215
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
         Left            =   510
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
         Left            =   495
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
Private oSearchEngine As z_SearchEngineB
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
      
      PopupMenu Forms(0).mnuBrowseSingles   ' Display the File menu as a
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
    oPC.OpenDBSHort
    '--------------
    oSearchEngine.SetupSQLwoCriteria2 False, pSearchType, False, CLng(txtmaxnum) + 1, strTypes  '"NGM"
    
    If pSearchType = enSearchByCatalogue Then
        oSearchEngine.selectcriteria "Catalogue", pCriteria, lngRecsFound
    ElseIf pSearchType = enSearchNormal Then
        oSearchEngine.SimpleSearch pCriteria, lngRecsFound
    Else
        enSource = enLocalDB
        If pSection <> "<ALL>" Then
            lngSectionID = oPC.Configuration.Sections.key(pSection)
        End If
        If pProductType <> "<ALL>" Then
            lngProductTypeID = oPC.Configuration.ProductTypes.key(pProductType)
        End If
        oSearchEngine.AdvancedSearch lngRecsFound, pCriteria, lngSectionID, lngProductTypeID
    End If
    'If lngRecsFound > CLng(txtmaxnum) Then MsgBox "Too many records to return, refine your search.", vbInformation + vbOKOnly, "Search result"
    oSearchEngine.executeSingle IIf(IsNumeric(txtmaxnum), CLng(txtmaxnum), 500)
    Set colList = Nothing
    Set colList = oSearchEngine.getcols
    lngrows = oSearchEngine.Rows
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
        XN.ReDim 1, colList.Count, 1, 12
        For i = 1 To colList.Count
                XN.Value(i, 1) = colList.Item(i).CodeF
                XN.Value(i, 2) = colList.Item(i).Title
                XN.Value(i, 3) = colList.Item(i).Length & "  " & colList.Item(i).Width
                XN.Value(i, 4) = colList.Item(i).LocalPriceF
                XN.Value(i, 9) = colList.Item(i).LocalPrice
                XN.Value(i, 10) = colList.Item(i).img
                XN.Value(i, 11) = colList.Item(i).PID
                XN.Value(i, 12) = colList.Item(i).EAN
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
    On Error GoTo errHandler
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
    Set oSearchEngine = New z_SearchEngineB
    Set colList = New Collection
    
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
        Me.top = 20
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
    

    GN.Columns(3).Caption = "Supplier"
    txtmaxnum = 500
    
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
        cboSearch(i).Text = tlCollection(i + 1).ItemByOrdinalIndex(1)
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
            strSQL = mDescriptionOrCode
        End If
    End If

    For i = 1 To 8
        If cboSearch(i - 1).Text <> "<ANY>" And cboSearch(i - 1).Text <> "" Then
            strSQL = strSQL & " PATINDEX('%" & tlCollection(i).key(cboSearch(i - 1).Text) & "%',dbo.FlattenCategorization(P_ID)) > 0 AND "
        End If
    Next
    If Right(strSQL, 5) = " AND " Then
        strSQL = Left(strSQL, Len(strSQL) - 5)
    End If

    If mLength > 0 Then
        strSQL = strSQL & " AND P_LENGTH < " & CStr(mLength) & " AND P_LENGTH > " & CStr(IIf(mLength > 500, mLength - 500, 0))
    End If
    If mWidth > 0 Then
        strSQL = strSQL & " AND P_WIDTH < " & CStr(mWidth) & " AND P_WIDTH > " & CStr(IIf(mWidth > 500, mWidth - 500, 0))
    End If
    
    If mPrice > 0 Then
        strSQL = strSQL & " AND P_SP < " & CStr(mPrice)
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
    On Error GoTo errHandler
Dim lngDiff As Long
    GN.Width = NonNegative_Lng(Me.Width - (GN.Left + 400))
    lngDiff = GN.Height
    GN.Height = NonNegative_Lng(Me.Height - (GN.top + 1220))
    lngDiff = (GN.Height - lngDiff)
    cmdclose.top = cmdclose.top + lngDiff

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
    On Error Resume Next
    Clipboard.Clear
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
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText Left(str, ISBNLENGTH)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.GN_RowColChange(LastRow,LastCol)", Array(LastRow, LastCol), _
         EA_NORERAISE
    HandleError
End Sub


Private Sub GN_DblClick()
    On Error GoTo errHandler
'Dim frmA As frmProductPrevAQ
'Dim frm As frmProductPrev
Dim frmNB As frmProductSinglePreview
Dim lngprod As Long
Dim str As String
    If XN.UpperBound(1) = 0 Then Exit Sub
    If IsNull(GN.Bookmark) Then Exit Sub
    If Err Then Exit Sub
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
    XA.ReDim 1, 1, 1, 7
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
    Cancel = Not IsNumeric(txtLength)
    mLength = ConvertDimensionsforStoring(FNDBL(txtLength))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtLength_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
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
    Cancel = Not IsNumeric(txtWidth)
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
    Cancel = Not IsNumeric(txtPrice)
    mPrice = FNDBL(txtPrice)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.txtPrice_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Public Sub mnuCreateInvoice()
    On Error GoTo errHandler
Dim frm As New frmPlacePF
Dim str As String
Dim top As Integer
Dim i As Integer

    If GN = "" Or GN = "No records" Then Exit Sub
    If GN.Bookmark = 0 Then Exit Sub

    'top = XA.UpperBound(1)
    XA.ReDim 1, XA.UpperBound(1) + GN.SelBookmarks.Count, 1, 9
    For i = 1 To GN.SelBookmarks.Count
        XA(i, 1) = FNS(XN.Value(GN.SelBookmarks(i - 1), 1))
        XA(i, 2) = FNS(XN.Value(GN.SelBookmarks(i - 1), 2))
        XA(i, 3) = FNS(XN.Value(GN.SelBookmarks(i - 1), 3))
        XA(i, 4) = 1
        XA(i, 5) = 0
        XA(i, 6) = ""
        XA(i, 7) = FNS(XN.Value(GN.SelBookmarks(i - 1), 9))
        XA(i, 8) = FNS(XN.Value(GN.SelBookmarks(i - 1), 11))
     '   XA(top + i, 9) = FNS(XN.Value(GN.SelBookmarks(i - 1), 16))
    Next

    frm.component XA, "INVOICE"
    frm.Show vbModal
    StartNewList
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmBrowseSingles.mnuCreateInvoice"
End Sub

