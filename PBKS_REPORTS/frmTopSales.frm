VERSION 5.00
Object = "{E6CC263E-5760-49D9-B793-4245D54496CF}#1.0#0"; "ExComboBox.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTopSales 
   BackColor       =   &H00D3D3CB&
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      Picture         =   "frmTopSales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2430
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3780
      Picture         =   "frmTopSales.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2430
      Width           =   1000
   End
   Begin VB.Frame fraPreviewPrint 
      BackColor       =   &H00D3D3CB&
      Height          =   1455
      Left            =   315
      TabIndex        =   3
      Top             =   2055
      Width           =   1665
      Begin VB.OptionButton optCSV 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&CSV"
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   150
         TabIndex        =   14
         Top             =   945
         Width           =   1065
      End
      Begin VB.OptionButton optPreview 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Pre&view"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   165
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton optPrint 
         BackColor       =   &H00D3D3CB&
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   555
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdSupp 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&per &supplier"
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
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   735
      Width           =   1230
   End
   Begin VB.TextBox txtSupplier 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1530
      TabIndex        =   1
      Top             =   750
      Width           =   2550
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00D3D3CB&
      Caption         =   "&All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   750
      Width           =   660
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1215
      TabIndex        =   6
      Top             =   105
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
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
      Format          =   191299585
      CurrentDate     =   37421
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   3390
      TabIndex        =   7
      Top             =   90
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
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
      Format          =   191299585
      CurrentDate     =   37421
   End
   Begin EXCOMBOBOXLibCtl.ComboBox cboProductType 
      Height          =   390
      Left            =   1545
      OleObjectBlob   =   "frmTopSales.frx":0714
      TabIndex        =   10
      Top             =   1410
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product type"
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
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   1440
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Height          =   270
      Left            =   2745
      TabIndex        =   9
      Top             =   135
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "between"
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
      Height          =   270
      Left            =   240
      TabIndex        =   8
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmTopSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim lngTPID As Long
Dim strSupplierName As String
Dim lngPTID As Long
Dim strPT As String

Public Sub Component(pCaption As String)
    Me.Caption = pCaption
End Sub
Private Sub SetupPT()
    cboProductType.BeginUpdate
    cboProductType.WidthList = 190
    cboProductType.HeightList = 162
    cboProductType.AllowSizeGrip = True
    cboProductType.AutoDropDown = True
    cboProductType.SelForeColor = vbRed
    cboProductType.Columns.Add "Product type"
    cboProductType.Columns.Add "Seesafe"
    cboProductType.Columns(0).Width = 190
    cboProductType.Columns(1).Width = 0
    cboProductType.BackColorLock = Me.BackColor
    cboProductType.EndUpdate
End Sub

Private Sub cmdAll_Click()
    strSupplierName = "<ALL>"
    lngTPID = 0
    txtSupplier = strSupplierName
End Sub

Private Sub cmdClose_Click()
    bCancelled = True
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub cmdSupp_Click()
Dim frm As frmBrowseSUppliers2
    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    lngTPID = frm.SupplierID
    strSupplierName = frm.SupplierName
    txtSupplier = strSupplierName
    Unload frm
    If lngTPID = 0 Then Exit Sub

End Sub


'Private Sub Form_Initialize()
'Dim ar() As String
'    cboProductType.BeginUpdate
'    oPC.Configuration.ProductTypes.CollectionAsArray ar
'    cboProductType.PutItems ar
'    cboProductType.EndUpdate
'
'End Sub

Private Sub Form_Load()
Dim ar() As String
    SetupPT
    cboProductType.BeginUpdate
    oPC.Configuration.ProductTypes.CollectionAsArray ar
    cboProductType.PutItems ar
    cboProductType.EndUpdate
    dtpFrom.Value = Date
    dtpTo.Value = Date
    Width = 5500
    Height = 4100
    bCancelled = False
End Sub


Private Sub cboProductType_SelectionChanged()
    lngPTID = oPC.Configuration.ProductTypes.Key(cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0))
    strPT = cboProductType.Items.CellCaption(cboProductType.Items.SelectedItem, 0)
End Sub

Property Get SupplierID() As Long
    SupplierID = lngTPID
End Property
Property Get PTID() As Long
    PTID = lngPTID
End Property
Property Get StartDate() As Date
    StartDate = CDate(dtpFrom.Value)
End Property
Property Get EndDate() As Date
    EndDate = CDate(dtpTo.Value)
End Property
Property Get SupplierName() As String
    SupplierName = strSupplierName
End Property
Property Get PTName() As String
    PTName = strPT
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

