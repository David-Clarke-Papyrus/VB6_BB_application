VERSION 5.00
Begin VB.Form frmCustomerReorderConfig 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Prepare ordering slate"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
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
      Left            =   660
      MaxLength       =   30
      TabIndex        =   9
      Top             =   5745
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Loading options"
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
      Height          =   1965
      Left            =   255
      TabIndex        =   7
      Top             =   1575
      Width           =   3090
      Begin VB.CheckBox chkActioned 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Exclude items which have already been actioned."
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   270
         TabIndex        =   8
         Top             =   945
         Value           =   1  'Checked
         Width           =   2550
      End
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
      Left            =   225
      TabIndex        =   4
      Top             =   3810
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
         TabIndex        =   6
         Top             =   330
         Width           =   2505
      End
      Begin VB.CheckBox chkOO 
         BackColor       =   &H00D3D3CB&
         Caption         =   "Filter only where P.O. = 0"
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
         TabIndex        =   5
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
      TabIndex        =   1
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
         Picture         =   "frmCustomerReorderConfig.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
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
         TabIndex        =   3
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
      Left            =   2370
      Picture         =   "frmCustomerReorderConfig.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5445
      Width           =   1000
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
      Left            =   735
      TabIndex        =   10
      Top             =   5490
      Width           =   1410
   End
End
Attribute VB_Name = "frmCustomerReorderConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frm As frmREORDER_CO
Dim bCancelled As Boolean
Dim mSupplierID As Long
Public Property Get Slatename() As String
    On Error GoTo errHandler
    Slatename = FNS(txtSlateName)
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.Slatename"
End Property
Public Sub component(pfrm As Form, pType As String)
    On Error GoTo errHandler
    Set frm = pfrm
    Me.txtSlateName = "Cust"
'   ' Me.DTPicker1 = Date
'    Me.frmSales.Enabled = Not (UCase(pType) = "CUST")
'    Option1.Enabled = Me.frmSales.Enabled
'    Option1.Value = IIf(Option1.Enabled, 1, 0)
'    Option2.Enabled = Me.frmSales.Enabled
'    Option3.Enabled = Me.frmSales.Enabled
'    Option4.Enabled = Me.frmSales.Enabled
'    Option5.Enabled = Me.frmSales.Enabled
'    Me.Label2.Enabled = Me.frmSales.Enabled
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.component(pfrm,pType)", Array(pfrm, pType)
End Sub



Private Sub cmdGo_Click()
    On Error GoTo errHandler
'   ' frm.Component "SALE"
'    If Me.Option1 Then
'        frm.Component2 DateAdd("ww", -1, Date), chkOH = 1, chkOO = 1, mSupplierID
'    ElseIf Option2 Then
'        frm.Component2 DateAdd("ww", -2, Date), chkOH = 1, chkOO = 1, mSupplierID
'    ElseIf Option3 Then
'        frm.Component2 DateAdd("ww", -3, Date), chkOH = 1, chkOO = 1, mSupplierID
'    ElseIf Option4 Then
'        frm.Component2 DateAdd("ww", -4, Date), chkOH = 1, chkOO = 1, mSupplierID
'    ElseIf Option5 Then
'        frm.Component2 Me.DTPicker1, chkOH = 1, chkOO = 1, mSupplierID
'    End If
    SaveSetting "PBKS", "frmCustomerReorderConfig", chkActioned, chkActioned

    'chkactioned
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.cmdGo_Click", , EA_NORERAISE
    HandleError
End Sub

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
    ErrorIn "frmCustomerReorderConfig.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    chkActioned = GetSetting("PBKS", "frmCustomerReorderConfig", chkActioned, 1)
    bCancelled = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.Form_Load", , EA_NORERAISE
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
    ErrorIn "frmCustomerReorderConfig.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode), _
         EA_NORERAISE
    HandleError
End Sub

Public Property Get Cancelled() As Boolean
    On Error GoTo errHandler
    Cancelled = bCancelled
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.Cancelled"
End Property

Public Property Get ExclActioned() As Boolean
    On Error GoTo errHandler
    ExclActioned = Me.chkActioned
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.ExclActioned"
End Property

Public Property Get OHFilter() As Boolean
    On Error GoTo errHandler
    OHFilter = Me.chkOH
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.OHFilter"
End Property
Public Property Get OOFilter() As Boolean
    On Error GoTo errHandler
    OOFilter = Me.chkOO
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.OOFilter"
End Property


Public Property Get SupplierID() As Long
    On Error GoTo errHandler
    SupplierID = mSupplierID
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.SupplierID"
End Property

Private Sub optActioned_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCustomerReorderConfig.optActioned_Click", , EA_NORERAISE
    HandleError
End Sub
