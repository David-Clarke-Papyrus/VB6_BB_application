VERSION 5.00
Begin VB.Form frmOrderFUlfil_Selection 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Selection for order fulfilment"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00D3D3CB&
      Caption         =   "5. or orders for stock from a supplier"
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
      Height          =   1170
      Left            =   240
      TabIndex        =   18
      Top             =   6375
      Width           =   3975
      Begin VB.CommandButton cmdSelectSupplier 
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
         Height          =   390
         Left            =   540
         Picture         =   "frmOrderFUlfil_Selection.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   630
         Width           =   630
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Choose supplier"
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
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   330
         Width           =   1530
      End
      Begin VB.Label lblSupplierName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   1800
         TabIndex        =   20
         Top             =   495
         Width           =   2085
      End
   End
   Begin VB.CheckBox chkPartial 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Exclude items which you can only partly fulfil."
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   315
      TabIndex        =   17
      Top             =   7920
      Width           =   2550
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00D3D3CB&
      Caption         =   "4. or by customer range"
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
      Height          =   1155
      Left            =   240
      TabIndex        =   14
      Top             =   5100
      Width           =   3975
      Begin VB.TextBox txtFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1410
         TabIndex        =   3
         Text            =   "a"
         Top             =   450
         Width           =   780
      End
      Begin VB.TextBox txtTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   2730
         TabIndex        =   4
         Text            =   "zzz"
         Top             =   450
         Width           =   780
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "names from"
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
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   465
         Width           =   1290
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Height          =   255
         Left            =   2265
         TabIndex        =   15
         Top             =   495
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "3. or orders for single customer"
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
      Height          =   1170
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   3975
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
         Height          =   390
         Left            =   540
         Picture         =   "frmOrderFUlfil_Selection.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   630
         Width           =   630
      End
      Begin VB.Label lblCustName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   1800
         TabIndex        =   13
         Top             =   495
         Width           =   2085
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Choose customer"
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
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   330
         Width           =   1530
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D3D3CB&
      Caption         =   "2. or orders captured by operator"
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
      Height          =   900
      Left            =   240
      TabIndex        =   9
      Top             =   2820
      Width           =   3975
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   2550
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   330
         Width           =   750
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Operator code"
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
         Height          =   285
         Left            =   990
         TabIndex        =   10
         Top             =   375
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D3D3CB&
      Caption         =   "1. Choose last set"
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
      Height          =   840
      Left            =   240
      TabIndex        =   0
      Top             =   1860
      Width           =   3975
      Begin VB.CheckBox chkLastSet 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D3D3CB&
         Caption         =   "Load last set"
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
         Height          =   450
         Left            =   2115
         TabIndex        =   8
         Top             =   255
         Width           =   1545
      End
   End
   Begin VB.CommandButton cmdFetch 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Load "
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
      Left            =   3255
      Picture         =   "frmOrderFUlfil_Selection.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   1000
   End
   Begin VB.Label lblOperator 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      Left            =   1500
      TabIndex        =   7
      Top             =   3255
      Width           =   1650
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmOrderFUlfil_Selection.frx":0A9E
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
      Height          =   1470
      Left            =   210
      TabIndex        =   6
      Top             =   195
      Width           =   4095
   End
End
Attribute VB_Name = "frmOrderFUlfil_Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOpCode As Long
Dim mCustFrom As String
Dim mCustTo As String
Dim mCustID As Long
Dim mSupplierID As Long
Dim mCustomerName As String
Dim bCancel As Boolean
Dim bComplete As Boolean
Dim bLoadLastSet As Boolean

Public Property Get CompleteOnly() As Boolean
    CompleteOnly = bComplete
End Property
Public Property Get LoadLastSet() As Boolean
    LoadLastSet = bLoadLastSet
End Property
Public Property Get CancelledYN() As Boolean
    CancelledYN = bCancel
End Property

Private Sub Check1_Click()
    On Error GoTo errHandler

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.Check1_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkLastSet_Click()
    On Error GoTo errHandler
    bLoadLastSet = (chkLastSet = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.chkLastSet_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub chkPartial_Click()
    On Error GoTo errHandler
    bComplete = (chkPartial = 1)
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.chkPartial_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdFetch_Click()
    On Error GoTo errHandler
    SaveSetting "PBKS", Me.Name, "Partial", CStr(chkPartial)
   
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.cmdFetch_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get oPCode() As Long
    oPCode = mOpCode
End Property
Public Property Get CustFrom() As String
    CustFrom = mCustFrom
End Property
Public Property Get CustTo() As String
    CustTo = mCustTo
End Property
Public Property Get CustID() As Long
    CustID = mCustID
End Property
Public Property Get SupplierID() As Long
    SupplierID = mSupplierID
End Property


Private Sub cmdLoadlast_Click()
    On Error GoTo errHandler
    bLoadLastSet = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.cmdLoadlast_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo errHandler
Dim frm As frmBrowseCustomers2
    Set frm = New frmBrowseCustomers2
    frm.Show vbModal
    mCustID = frm.CustomerID
    Me.lblCustName = frm.CustomerName
    Unload frm
    If mCustID = 0 Then Exit Sub
    txtFrom = ""
    txtTo = ""
    mCustTo = ""
   mCustFrom = ""

    txtPassword = ""
    lblOperator.Caption = ""
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.cmdSelect_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub cmdSelectSupplier_Click()
    On Error GoTo errHandler
Dim frm As frmBrowseSUppliers2
    Set frm = New frmBrowseSUppliers2
    frm.Show vbModal
    mSupplierID = frm.SupplierID
    Me.lblSupplierName = frm.SupplierName
    Unload frm
    If mSupplierID = 0 Then Exit Sub
    txtFrom = ""
    txtTo = ""
    mCustTo = ""
   mCustFrom = ""

    txtPassword = ""
    lblOperator.Caption = ""

    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.cmdSelectSupplier_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
  mCustTo = "zzz"
  mCustFrom = "a"
  bCancel = False
  bLoadLastSet = False
  chkPartial = val(GetSetting("PBKS", Me.Name, "Partial", 0))
 
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.Form_Load", , EA_NORERAISE
    HandleError
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo errHandler
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        bCancel = True
        Me.Hide
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.Form_QueryUnload(Cancel,UnloadMode)", Array(Cancel, UnloadMode), _
         EA_NORERAISE
    HandleError
End Sub

Private Sub txtFrom_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtFrom
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.txtFrom_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtPassword_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtPassword
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.txtPassword_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTo_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtTo
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.txtTo_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtTo_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    mCustTo = txtTo
    If mCustFrom > "" Then
        mCustID = 0
        txtPassword = ""
        lblCustName.Caption = ""
        lblOperator.Caption = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.txtTo_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtFrom_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    mCustFrom = txtFrom
    If mCustTo > "" Then
        mCustID = 0
        txtPassword = ""
        lblCustName.Caption = ""
        lblOperator.Caption = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.txtFrom_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub

Private Sub txtPassword_Validate(Cancel As Boolean)
    On Error GoTo errHandler
Dim lngStaffID As Long
Dim strName As String

    oPC.Configuration.Staff.GetLevel txtPassword, strName, lngStaffID
    mOpCode = lngStaffID
    lblOperator.Caption = strName
    If mOpCode > 0 Then
        mCustID = 0
        txtFrom = ""
        txtTo = ""
        Me.lblCustName.Caption = ""
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOrderFUlfil_Selection.txtPassword_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
