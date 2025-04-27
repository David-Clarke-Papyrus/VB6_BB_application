VERSION 5.00
Begin VB.Form frmCatNo 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Catalogue label printing"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3360
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C4BCA4&
      Cancel          =   -1  'True
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
      Height          =   495
      Left            =   1635
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click to find all customers matching the retrictions selected."
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   750
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C4BCA4&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   885
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Click to find all customers matching the retrictions selected."
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   750
   End
   Begin VB.TextBox txtCatNo 
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
      Height          =   330
      Left            =   795
      TabIndex        =   0
      ToolTipText     =   "Customer name starts like this"
      Top             =   510
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00D3D3CB&
      Caption         =   "Catalogue number"
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
      Left            =   825
      TabIndex        =   1
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "frmCatNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancel As Boolean
Dim strCatNo As String

Private Sub cmdCancel_Click()
    On Error GoTo errHandler
    bCancel = True
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatNo.cmdCancel_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandler
    bCancel = False
    strCatNo = Me.txtCatNo
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmCatNo.cmdOK_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get Cancelled() As Boolean
    Cancelled = bCancel
End Property

Public Property Get CatNo() As String
    CatNo = strCatNo
End Property
