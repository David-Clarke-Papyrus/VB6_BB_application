VERSION 5.00
Begin VB.Form frmDam 
   Caption         =   "Damaged items"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   2625
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   15
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmDam.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1485
      Width           =   1110
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   945
      TabIndex        =   1
      Top             =   555
      Width           =   630
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
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
      Height          =   675
      Left            =   1500
      MaskColor       =   &H00C4BCA4&
      Picture         =   "frmDam.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1485
      Width           =   1110
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Must be numeric"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   645
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00D3D3CB&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty damaged"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   255
      TabIndex        =   2
      Top             =   240
      Width           =   2100
   End
End
Attribute VB_Name = "frmDam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean

Public Sub component(pQty As Long)
    On Error GoTo errHandler
    txtQty = pQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDam.component(pQty)", pQty
End Sub

Public Property Get DamagedQty() As String
    DamagedQty = txtQty
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = bCancelled
End Property

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancelled = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDam.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDam.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub



Private Sub txtQty_Change()
    On Error GoTo errHandler
    lblMsg.Visible = Not (IsNumeric(txtQty))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDam.txtQty_Change", , EA_NORERAISE
    HandleError
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    Cancel = Not (IsNumeric(txtQty))
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmDam.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
