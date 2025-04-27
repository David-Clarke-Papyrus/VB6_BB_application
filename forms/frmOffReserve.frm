VERSION 5.00
Begin VB.Form frmOffReserve 
   BackColor       =   &H00D3D3CB&
   Caption         =   "Collect from reserve"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDetails 
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
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1410
      Width           =   3135
   End
   Begin VB.TextBox txtQty 
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
      Left            =   1590
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   405
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2115
      Width           =   3150
   End
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00D7D1BF&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1170
   End
   Begin VB.CommandButton cmdYes 
      BackColor       =   &H00D7D1BF&
      Caption         =   "Yes"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   660
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2865
      Width           =   1170
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copies have been collected and must be taken off reserve."
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
      Height          =   735
      Left            =   630
      TabIndex        =   5
      Top             =   645
      Width           =   2700
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
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
      Height          =   210
      Left            =   405
      TabIndex        =   4
      Top             =   1875
      Width           =   705
   End
End
Attribute VB_Name = "frmOffReserve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strChoice As String
Dim strMsg As String
Dim lngQty As Long


Private Sub cmdNo_Click()
    On Error GoTo errHandler
    strChoice = "No"
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserve.cmdNo_Click", , EA_NORERAISE
    HandleError
End Sub

Private Sub cmdYes_Click()
    On Error GoTo errHandler
    strChoice = "Yes"
    Me.Hide
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserve.cmdYes_Click", , EA_NORERAISE
    HandleError
End Sub

Public Property Get GetChoice() As String
    GetChoice = strChoice
End Property
Public Property Get GetNote() As String
    GetNote = strMsg
End Property

Private Sub txtNote_Change()
    On Error GoTo errHandler
    strMsg = txtNote
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserve.txtNote_Change", , EA_NORERAISE
    HandleError
End Sub

Public Sub component(pDetails As String, pQty As Long)
    On Error GoTo errHandler
    txtDetails = pDetails
    lngQty = pQty
    txtQty = lngQty
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserve.component(pDetails,pQty)", Array(pDetails, pQty)
End Sub

Public Property Get GetQty() As Long
    On Error GoTo errHandler
    GetQty = lngQty
    Exit Property
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserve.GetQty"
End Property

Private Sub txtQty_Validate(Cancel As Boolean)
    On Error GoTo errHandler
    If Not ConvertToLng(txtQty, lngQty) Then
        Cancel = True
    End If
    Exit Sub
errHandler:
    If ErrMustStop Then Debug.Assert False: Resume
    ErrorIn "frmOffReserve.txtQty_Validate(Cancel)", Cancel, EA_NORERAISE
    HandleError
End Sub
