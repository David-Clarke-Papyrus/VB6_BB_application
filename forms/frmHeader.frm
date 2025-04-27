VERSION 5.00
Begin VB.Form frmHeader 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Quotation header details"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5310
   Begin VB.TextBox txtForAttn 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1245
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1125
      Width           =   3870
   End
   Begin VB.TextBox txtOrderDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2475
      MaxLength       =   20
      TabIndex        =   1
      Top             =   3015
      Width           =   1560
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "C&ancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6195
      Width           =   135
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&Close"
      Default         =   -1  'True
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
      Left            =   4080
      Picture         =   "frmHeader.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1000
   End
   Begin VB.TextBox txtMemo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   630
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   165
      Width           =   4485
   End
   Begin VB.TextBox txtOrderNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2490
      MaxLength       =   20
      TabIndex        =   0
      Top             =   2580
      Width           =   2625
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "For attn."
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
      Height          =   240
      Left            =   195
      TabIndex        =   10
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label lblOrderDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Order date"
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
      Height          =   240
      Left            =   1425
      TabIndex        =   8
      Top             =   3030
      Width           =   900
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "(Click ESC to cancel)"
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   75
      TabIndex        =   7
      Top             =   1905
      Width           =   1800
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Memo"
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
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   150
      Width           =   450
   End
   Begin VB.Label lblOrderNumber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Request for quote reference"
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
      Height          =   240
      Left            =   150
      TabIndex        =   4
      Top             =   2595
      Width           =   2175
   End
End
Attribute VB_Name = "frmHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strOrderNumber As String
Dim strOrderDate As String
Dim strMemo As String
Dim flgLoading As Boolean
Dim oCO As a_CO
Dim bCancel As Boolean
Dim Blocked As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Sub component()
    On Error GoTo errHandler

    If Me.WindowState <> 2 Then
        TOP = 3000
        Left = 1000
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader.component"
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdClose_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader.cmdClose_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub Form_Load()
    On Error GoTo errHandler
    bCancel = False
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader.Form_Load", , EA_NORERAISE
    HandleError
End Sub




'Private Sub txtOrderDate_GotFocus()
'    AutoSelect txtOrderDate
'End Sub
'
'Private Sub txtOrderDate_LostFocus()
'    Me.cmdClose.Enabled = (IsDate(txtOrderDate) Or txtOrderDate = "")
'End Sub
'
'Private Sub txtOrderNumber_GotFocus()
'    AutoSelect txtOrderNumber
'End Sub
'
'Public Property Get OrderNumber() As String
'    OrderNumber = txtOrderNumber
'End Property
'
'Public Property Get OrderDate() As String
'    OrderDate = txtOrderDate
'End Property
Public Property Get Memo() As String
    Memo = txtMemo
End Property

Private Sub txtForAttn_GotFocus()
    On Error GoTo errHandler
    AutoSelect txtForAttn
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmHeader.txtForAttn_GotFocus", , EA_NORERAISE
    HandleError
End Sub

Public Property Get ForAttn() As String
    ForAttn = txtForAttn
End Property

