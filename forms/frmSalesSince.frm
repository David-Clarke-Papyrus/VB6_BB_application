VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesSince 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Prepare ordering slate"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   3690
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSlateName 
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
      Height          =   315
      Left            =   645
      MaxLength       =   30
      TabIndex        =   15
      Top             =   6870
      Width           =   1545
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Order and O.H. filter"
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
      Height          =   1275
      Left            =   240
      TabIndex        =   12
      Top             =   4710
      Width           =   3135
      Begin VB.CheckBox chkOH 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Filter only where O.H. = 0"
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
         Height          =   315
         Left            =   300
         TabIndex        =   14
         Top             =   330
         Width           =   2505
      End
      Begin VB.CheckBox chkOO 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Filter only where O.O. = 0"
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
         Height          =   315
         Left            =   300
         TabIndex        =   13
         Top             =   690
         Width           =   2505
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Supplier filter"
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
      Height          =   1155
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   3105
      Begin VB.CommandButton cmdSelect 
         BackColor       =   &H00C4BCA4&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2400
         Picture         =   "frmSalesSince.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lblSupplierName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "                                                <All suppliers>"
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
         Height          =   570
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   2115
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdGo 
      BackColor       =   &H00C4BCA4&
      Caption         =   "Fetch"
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
      Left            =   2340
      Picture         =   "frmSalesSince.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6585
      Width           =   1000
   End
   Begin VB.Frame frmSales 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Select sales made over last"
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
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   1380
      Width           =   3120
      Begin VB.OptionButton Option5 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Or"
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
         Left            =   180
         TabIndex        =   8
         Top             =   2160
         Width           =   600
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00D3D3CB&
         Caption         =   "4 weeks"
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
         Left            =   855
         TabIndex        =   4
         Top             =   1680
         Width           =   1260
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00D3D3CB&
         Caption         =   "3 weeks"
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
         Left            =   855
         TabIndex        =   3
         Top             =   1230
         Width           =   1260
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D3D3CB&
         Caption         =   "2 weeks"
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
         Left            =   855
         TabIndex        =   2
         Top             =   810
         Width           =   1260
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D3D3CB&
         Caption         =   "1 week"
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
         Left            =   855
         TabIndex        =   1
         Top             =   375
         Value           =   -1  'True
         Width           =   1260
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   1470
         TabIndex        =   6
         Top             =   2400
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   221839361
         CurrentDate     =   38987
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Since"
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
         Left            =   810
         TabIndex        =   7
         Top             =   2220
         Width           =   555
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Slate name"
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
      Left            =   720
      TabIndex        =   16
      Top             =   6615
      Width           =   1410
   End
End
Attribute VB_Name = "frmSalesSince"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frm As frmREORDER_CO
Dim bCancelled As Boolean
Dim mSupplierID As Long

Public Sub component(pfrm As Form, pType As String)
    On Error GoTo errHandler
    Set frm = pfrm
    Me.DTPicker1 = Date
    Me.frmSales.Enabled = Not (UCase(pType) = "CUST")
    Option1.Enabled = Me.frmSales.Enabled
    Option1.Value = IIf(Option1.Enabled, 1, 0)
    Option2.Enabled = Me.frmSales.Enabled
    Option3.Enabled = Me.frmSales.Enabled
    Option4.Enabled = Me.frmSales.Enabled
    Option5.Enabled = Me.frmSales.Enabled
    Me.Label2.Enabled = Me.frmSales.Enabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.component(pfrm,pType)", Array(pfrm, pType)
End Sub

Public Property Get OHFilter() As Boolean
    On Error GoTo errHandler
    OHFilter = Me.chkOH
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.OHFilter"
End Property
Public Property Get OOFilter() As Boolean
    On Error GoTo errHandler
    OOFilter = Me.chkOO
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.OOFilter"
End Property
Private Sub cmdGo_Click()
    On Error GoTo errHandler
   ' frm.Component "SALE"
    If Me.Option1 Then
        frm.Component2 DateAdd("ww", -1, Date), chkOH = 1, chkOO = 1, mSupplierID
    ElseIf Option2 Then
        frm.Component2 DateAdd("ww", -2, Date), chkOH = 1, chkOO = 1, mSupplierID
    ElseIf Option3 Then
        frm.Component2 DateAdd("ww", -3, Date), chkOH = 1, chkOO = 1, mSupplierID
    ElseIf Option4 Then
        frm.Component2 DateAdd("ww", -4, Date), chkOH = 1, chkOO = 1, mSupplierID
    ElseIf Option5 Then
        frm.Component2 Me.DTPicker1, chkOH = 1, chkOO = 1, mSupplierID
    End If
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub
Public Property Get Slatename() As String
    On Error GoTo errHandler
    Slatename = FNS(txtSlateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Slatename"
End Property
Private Sub cmdSelect_Click()
    On Error GoTo errHandler
Dim frm As frmBrowseSUppliers2
    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    mSupplierID = frm.SupplierID
    Me.lblSupplierName = frm.SupplierName
    Unload frm
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    txtSlateName = Format(Now(), "YY-MM-DD HH:NN")
    bCancelled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Form_Load", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    If UnloadMode = vbFormControlMenu Then
        bCancelled = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode), _
         EA_NORERAISE
    HandleError
End Sub

Public Property Get Cancelled() As Boolean
    On Error GoTo errHandler
    Cancelled = bCancelled
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Cancelled"
End Property

Private Sub Option1_Click()
    On Error GoTo errHandler
    DTPicker1.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Option1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Option2_Click()
    On Error GoTo errHandler
    DTPicker1.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Option2_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Option3_Click()
    On Error GoTo errHandler
    DTPicker1.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Option3_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Option4_Click()
    On Error GoTo errHandler
    DTPicker1.Enabled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Option4_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Option5_Click()
    On Error GoTo errHandler
    DTPicker1.Enabled = True
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.Option5_Click", , EA_NORERAISE
    HandleError
End Sub
Public Property Get SupplierID() As Long
    On Error GoTo errHandler
    SupplierID = mSupplierID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmSalesSince.SupplierID"
End Property

